// mod.rs — MathML-to-OMML converter
//
// This module parses MathML XML using `quick_xml::Reader` and builds a tree
// of `OmmlNode` values that can be serialised to Office MathML.
//
// Ported from the Python `mathml2omml` package.

pub mod elements;
pub mod op_dict;

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use elements::{make_run_style, OmmlNode, RunKind, RunStyle};
use op_dict::{op_lookup, OpAttrs as DictOpAttrs, OpForm};

// ---------------------------------------------------------------------------
// Error type
// ---------------------------------------------------------------------------

/// Errors that can occur during MathML-to-OMML conversion.
#[derive(Debug)]
pub enum MathError {
    /// The input XML was malformed or could not be parsed.
    Xml(String),
    /// An unsupported MathML element was encountered.
    Unsupported(String),
    /// A structural error in the MathML (e.g. wrong number of children).
    Structure(String),
    /// An error from the LaTeX-to-MathML conversion step.
    Latex(String),
}

impl std::fmt::Display for MathError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            MathError::Xml(msg) => write!(f, "XML error: {}", msg),
            MathError::Unsupported(msg) => write!(f, "unsupported element: {}", msg),
            MathError::Structure(msg) => write!(f, "structure error: {}", msg),
            MathError::Latex(msg) => write!(f, "LaTeX error: {}", msg),
        }
    }
}

impl std::error::Error for MathError {}

impl From<quick_xml::Error> for MathError {
    fn from(e: quick_xml::Error) -> Self {
        MathError::Xml(e.to_string())
    }
}

// ---------------------------------------------------------------------------
// Resolved operator attributes (after dictionary lookup + element overrides)
// ---------------------------------------------------------------------------

/// Resolved operator attributes, combining dictionary defaults with any
/// explicit attribute overrides on the MathML element.
#[derive(Debug, Clone)]
struct ResolvedOpAttrs {
    fence: bool,
    stretchy: bool,
    accent: bool,
    largeop: bool,
    movablelimits: bool,
    form: &'static str,
}

/// Look up operator attributes, trying forms in priority order and applying
/// explicit overrides from the element's attributes.
///
/// Mirrors the Python `op_attrs` function.
fn resolve_op_attrs(
    content: &str,
    recommended_form: &str,
    node_attrs: &HashMap<String, String>,
) -> ResolvedOpAttrs {
    let forms_to_try: &[&str] = &[recommended_form, "infix", "prefix", "postfix"];

    let mut dict_entry: Option<&DictOpAttrs> = None;
    let mut found_form: &str = "infix";

    for &form_str in forms_to_try {
        if let Some(of) = OpForm::from_str(form_str) {
            if let Some(entry) = op_lookup(content, of) {
                dict_entry = Some(entry);
                found_form = match of {
                    OpForm::Prefix => "prefix",
                    OpForm::Infix => "infix",
                    OpForm::Postfix => "postfix",
                };
                break;
            }
        }
    }

    let mut attrs = ResolvedOpAttrs {
        fence: dict_entry.map_or(false, |e| e.fence),
        stretchy: dict_entry.map_or(false, |e| e.stretchy),
        accent: dict_entry.map_or(false, |e| e.accent),
        largeop: dict_entry.map_or(false, |e| e.largeop),
        movablelimits: dict_entry.map_or(false, |e| e.movablelimits),
        form: found_form,
    };

    // Apply explicit overrides from element attributes
    fn try_bool(val: &str) -> Option<bool> {
        match val {
            "true" => Some(true),
            "false" => Some(false),
            _ => None,
        }
    }
    if let Some(v) = node_attrs.get("fence").and_then(|v| try_bool(v)) {
        attrs.fence = v;
    }
    if let Some(v) = node_attrs.get("stretchy").and_then(|v| try_bool(v)) {
        attrs.stretchy = v;
    }
    if let Some(v) = node_attrs.get("accent").and_then(|v| try_bool(v)) {
        attrs.accent = v;
    }
    if let Some(v) = node_attrs.get("largeop").and_then(|v| try_bool(v)) {
        attrs.largeop = v;
    }
    if let Some(v) = node_attrs.get("movablelimits").and_then(|v| try_bool(v)) {
        attrs.movablelimits = v;
    }
    if let Some(v) = node_attrs.get("form") {
        // Re-lookup with the explicit form
        if let Some(of) = OpForm::from_str(v) {
            if let Some(entry) = op_lookup(content, of) {
                // Keep the explicit form but re-derive defaults from the dict
                if node_attrs.get("fence").is_none() {
                    attrs.fence = entry.fence;
                }
                if node_attrs.get("stretchy").is_none() {
                    attrs.stretchy = entry.stretchy;
                }
                if node_attrs.get("accent").is_none() {
                    attrs.accent = entry.accent;
                }
                if node_attrs.get("largeop").is_none() {
                    attrs.largeop = entry.largeop;
                }
                if node_attrs.get("movablelimits").is_none() {
                    attrs.movablelimits = entry.movablelimits;
                }
            }
            attrs.form = match of {
                OpForm::Prefix => "prefix",
                OpForm::Infix => "infix",
                OpForm::Postfix => "postfix",
            };
        }
    }

    attrs
}

/// Check whether an operator node is a stretchy accent.
fn is_stretch_accent(node: &OmmlNode) -> bool {
    if let OmmlNode::Run {
        kind: RunKind::Mo,
        text,
        attrs,
        ..
    } = node
    {
        let oa = resolve_op_attrs(text.trim(), "infix", attrs);
        oa.stretchy && oa.accent
    } else {
        false
    }
}

// ---------------------------------------------------------------------------
// Fence-merging (merge_mrow_elems)
// ---------------------------------------------------------------------------

/// Pending state used during fence merging (mirrors Python `Pending`).
enum PendingItem {
    /// A list of nodes being accumulated (the bottom of the stack).
    Nodes(Vec<OmmlNode>),
    /// An N-ary operator waiting for its content element.
    Nary(OmmlNode, Vec<OmmlNode>),
    /// An opening fence waiting for a closing fence.
    Fence(HashMap<String, String>, Vec<OmmlNode>),
}

fn pop_opening_fence(stack: &mut Vec<PendingItem>, close_fence: &str) {
    while stack.len() > 1 {
        let top = stack.pop().unwrap();
        match top {
            PendingItem::Fence(mut attrs, elems) => {
                attrs.insert("close".to_string(), close_fence.to_string());
                let open = attrs.get("open").cloned().unwrap_or_default();
                let close = attrs.get("close").cloned().unwrap_or_default();
                let seps = attrs.get("separators").cloned().unwrap_or_default();
                let node = OmmlNode::Delimited {
                    open,
                    close,
                    separators: seps,
                    children: elems,
                };
                push_to_top(stack, node);
                return;
            }
            PendingItem::Nary(base, elems) => {
                let fixed = fix_pending_nary(base, elems);
                push_to_top(stack, fixed);
                continue;
            }
            PendingItem::Nodes(_) => {
                unreachable!("Nodes should only be at the bottom of the stack");
            }
        }
    }
    // No matching open fence found — synthesise one wrapping everything so far
    let all = match &mut stack[0] {
        PendingItem::Nodes(ref mut v) => std::mem::take(v),
        _ => vec![],
    };
    let node = OmmlNode::Delimited {
        open: String::new(),
        close: close_fence.to_string(),
        separators: String::new(),
        children: all,
    };
    stack[0] = PendingItem::Nodes(vec![node]);
}

fn push_to_top(stack: &mut Vec<PendingItem>, node: OmmlNode) {
    match stack.last_mut() {
        Some(PendingItem::Nodes(v)) => v.push(node),
        Some(PendingItem::Nary(_, v)) => v.push(node),
        Some(PendingItem::Fence(_, v)) => v.push(node),
        None => {}
    }
}

fn fix_pending_nary(base: OmmlNode, elems: Vec<OmmlNode>) -> OmmlNode {
    // Try to extract the base MO operator from a script element and promote
    // to Nary. The base is an N-ary-able element (Sub/Sup/SubSup/Under/
    // Over/UnderOver) whose first child is a large prefix operator.
    macro_rules! try_nary {
        ($inner_base:expr, $sub:expr, $sup:expr, $default_loc:expr) => {
            if let OmmlNode::Run {
                kind: RunKind::Mo,
                ref text,
                ref attrs,
                ..
            } = *$inner_base
            {
                let oa = resolve_op_attrs(text.trim(), "prefix", attrs);
                let lim_loc = if oa.movablelimits {
                    "undOvr"
                } else {
                    $default_loc
                };
                return OmmlNode::Nary {
                    chr: text.clone(),
                    sub: $sub,
                    sup: $sup,
                    content: Box::new(row_or_single(elems)),
                    lim_loc: lim_loc.to_string(),
                };
            }
        };
    }

    match base {
        OmmlNode::Sub {
            base: inner_base,
            sub,
        } => {
            try_nary!(inner_base, sub, Box::new(OmmlNode::Empty), "subSup");
            let mut result = vec![OmmlNode::Sub {
                base: inner_base,
                sub,
            }];
            result.extend(elems);
            row_or_single(result)
        }
        OmmlNode::Sup {
            base: inner_base,
            sup,
        } => {
            try_nary!(inner_base, Box::new(OmmlNode::Empty), sup, "subSup");
            let mut result = vec![OmmlNode::Sup {
                base: inner_base,
                sup,
            }];
            result.extend(elems);
            row_or_single(result)
        }
        OmmlNode::SubSup {
            base: inner_base,
            sub,
            sup,
        } => {
            try_nary!(inner_base, sub, sup, "subSup");
            let mut result = vec![OmmlNode::SubSup {
                base: inner_base,
                sub,
                sup,
            }];
            result.extend(elems);
            row_or_single(result)
        }
        OmmlNode::Under {
            base: inner_base,
            under,
        } => {
            try_nary!(inner_base, under, Box::new(OmmlNode::Empty), "undOvr");
            let mut result = vec![OmmlNode::Under {
                base: inner_base,
                under,
            }];
            result.extend(elems);
            row_or_single(result)
        }
        OmmlNode::Over {
            base: inner_base,
            over,
        } => {
            try_nary!(inner_base, Box::new(OmmlNode::Empty), over, "undOvr");
            let mut result = vec![OmmlNode::Over {
                base: inner_base,
                over,
            }];
            result.extend(elems);
            row_or_single(result)
        }
        OmmlNode::UnderOver {
            base: inner_base,
            under,
            over,
        } => {
            try_nary!(inner_base, under, over, "undOvr");
            let mut result = vec![OmmlNode::UnderOver {
                base: inner_base,
                under,
                over,
            }];
            result.extend(elems);
            row_or_single(result)
        }
        other => {
            let mut result = vec![other];
            result.extend(elems);
            row_or_single(result)
        }
    }
}

fn row_or_single(mut nodes: Vec<OmmlNode>) -> OmmlNode {
    if nodes.len() == 1 {
        nodes.pop().unwrap()
    } else {
        OmmlNode::Row { children: nodes }
    }
}

/// Check whether a node is an "N-ary-able" element whose base child is a
/// large operator in prefix form.
fn is_nary(node: &OmmlNode) -> bool {
    let base = match node {
        OmmlNode::Sub { base, .. }
        | OmmlNode::Sup { base, .. }
        | OmmlNode::SubSup { base, .. }
        | OmmlNode::Under { base, .. }
        | OmmlNode::Over { base, .. }
        | OmmlNode::UnderOver { base, .. } => base.as_ref(),
        _ => return false,
    };
    if let OmmlNode::Run {
        kind: RunKind::Mo,
        text,
        attrs,
        ..
    } = base
    {
        let oa = resolve_op_attrs(text.trim(), "prefix", attrs);
        oa.largeop && oa.form == "prefix"
    } else {
        false
    }
}

fn append_elem(stack: &mut Vec<PendingItem>, elem: OmmlNode, recommended_form: &str) {
    // N-ary promotion
    if is_nary(&elem) {
        stack.push(PendingItem::Nary(elem, Vec::new()));
        return;
    }

    // Fence detection
    if let Some(core_op) = elem.embellished_op() {
        if let OmmlNode::Run {
            kind: RunKind::Mo,
            text,
            attrs,
            ..
        } = core_op
        {
            let oa = resolve_op_attrs(text.trim(), recommended_form, attrs);
            if oa.fence && oa.stretchy {
                match oa.form {
                    "prefix" => {
                        let mut fence_attrs = HashMap::new();
                        fence_attrs.insert("open".to_string(), text.clone());
                        fence_attrs.insert("close".to_string(), String::new());
                        fence_attrs.insert("separators".to_string(), String::new());
                        stack.push(PendingItem::Fence(fence_attrs, Vec::new()));
                        return;
                    }
                    "postfix" => {
                        pop_opening_fence(stack, text);
                        return;
                    }
                    "infix" => {
                        let has_open_fence = stack
                            .iter()
                            .rev()
                            .any(|item| matches!(item, PendingItem::Fence(..)));
                        if has_open_fence {
                            pop_opening_fence(stack, text);
                            return;
                        }
                        let mut fence_attrs = HashMap::new();
                        fence_attrs.insert("open".to_string(), text.clone());
                        fence_attrs.insert("close".to_string(), String::new());
                        fence_attrs.insert("separators".to_string(), String::new());
                        stack.push(PendingItem::Fence(fence_attrs, Vec::new()));
                        return;
                    }
                    _ => {}
                }
            }
        }
    }

    push_to_top(stack, elem);
}

/// Merge mrow children by detecting fence pairs and N-ary operators.
///
/// This mirrors the Python `merge_mrow_elems` function.
pub fn merge_mrow_elems(elements: Vec<OmmlNode>) -> Vec<OmmlNode> {
    if elements.is_empty() {
        return elements;
    }

    let len = elements.len();

    // Find the last non-space-like index
    let mut last_idx = len - 1;
    let mut found_last = false;
    for idx in (0..len).rev() {
        if !elements[idx].is_space_like() {
            last_idx = idx;
            found_last = true;
            break;
        }
    }
    if !found_last {
        return elements;
    }

    // Find the first non-space-like index
    let mut first_idx = 0;
    let mut found_first = false;
    for idx in 0..=last_idx {
        if !elements[idx].is_space_like() {
            first_idx = idx;
            found_first = true;
            break;
        }
    }
    if !found_first {
        return elements;
    }

    // Split into prefix-space, active-range, suffix-space
    let prefix: Vec<OmmlNode> = elements[..first_idx].to_vec();
    let suffix: Vec<OmmlNode> = if last_idx + 1 < len {
        elements[last_idx + 1..].to_vec()
    } else {
        vec![]
    };
    let active: Vec<OmmlNode> = elements[first_idx..=last_idx].to_vec();

    if active.is_empty() {
        return elements;
    }

    let mut stack: Vec<PendingItem> = vec![PendingItem::Nodes(Vec::new())];

    // First element: try as prefix
    let mut iter = active.into_iter();
    if let Some(first) = iter.next() {
        append_elem(&mut stack, first, "prefix");
    }

    // Remaining elements: middle ones as infix, last as postfix
    let rest: Vec<OmmlNode> = iter.collect();
    if !rest.is_empty() {
        let last_rest_idx = rest.len() - 1;
        for (i, elem) in rest.into_iter().enumerate() {
            if i < last_rest_idx {
                append_elem(&mut stack, elem, "infix");
            } else {
                append_elem(&mut stack, elem, "postfix");
            }
        }
    }

    // Collapse the stack
    while stack.len() > 1 {
        let top = stack.pop().unwrap();
        match top {
            PendingItem::Nary(base, elems) => {
                let fixed = fix_pending_nary(base, elems);
                push_to_top(&mut stack, fixed);
            }
            PendingItem::Fence(attrs, elems) => {
                let open = attrs.get("open").cloned().unwrap_or_default();
                let close = attrs.get("close").cloned().unwrap_or_default();
                let seps = attrs.get("separators").cloned().unwrap_or_default();
                let node = OmmlNode::Delimited {
                    open,
                    close,
                    separators: seps,
                    children: elems,
                };
                push_to_top(&mut stack, node);
            }
            PendingItem::Nodes(_) => unreachable!(),
        }
    }

    let mut result = prefix;
    if let Some(PendingItem::Nodes(nodes)) = stack.pop() {
        result.extend(nodes);
    }
    result.extend(suffix);
    result
}

// ---------------------------------------------------------------------------
// Builder — tracks stack state while parsing MathML events
// ---------------------------------------------------------------------------

/// Frame on the builder stack, representing an open MathML element.
enum Frame {
    /// `<math>` — top-level
    Math(Vec<OmmlNode>, HashMap<String, String>),
    /// `<mrow>`
    Row(Vec<OmmlNode>, HashMap<String, String>),
    /// Token element: mi, mn, mo, mtext, mspace, ms
    Token(RunKind, String, HashMap<String, String>),
    /// `<mfrac>`
    Frac(Vec<OmmlNode>, HashMap<String, String>),
    /// `<msqrt>`
    Sqrt(Vec<OmmlNode>, HashMap<String, String>),
    /// `<mroot>`
    Root(Vec<OmmlNode>, HashMap<String, String>),
    /// `<msub>`
    Sub(Vec<OmmlNode>, HashMap<String, String>),
    /// `<msup>`
    Sup(Vec<OmmlNode>, HashMap<String, String>),
    /// `<msubsup>`
    SubSup(Vec<OmmlNode>, HashMap<String, String>),
    /// `<munder>`
    Under(Vec<OmmlNode>, HashMap<String, String>),
    /// `<mover>`
    Over(Vec<OmmlNode>, HashMap<String, String>),
    /// `<munderover>`
    UnderOver(Vec<OmmlNode>, HashMap<String, String>),
    /// `<mfenced>`
    Fenced(Vec<OmmlNode>, HashMap<String, String>),
    /// `<mstyle>`
    Style(Vec<OmmlNode>, HashMap<String, String>),
    /// `<mpadded>`
    Padded(Vec<OmmlNode>, HashMap<String, String>),
    /// `<mphantom>`
    Phantom(Vec<OmmlNode>, HashMap<String, String>),
    /// `<menclose>`
    Enclose(Vec<OmmlNode>, HashMap<String, String>),
    /// `<mtable>`
    Table(Vec<Vec<OmmlNode>>, HashMap<String, String>),
    /// `<mtr>` or `<mlabeledtr>`
    TableRow(Vec<OmmlNode>, HashMap<String, String>),
    /// `<mtd>`
    TableData(Vec<OmmlNode>, HashMap<String, String>),
    /// `<maction>`
    Action(Vec<OmmlNode>, HashMap<String, String>),
    /// `<mmultiscripts>`
    MultiScripts(Vec<OmmlNode>, Option<Vec<OmmlNode>>, HashMap<String, String>),
    /// Ignored element (annotation, annotation-xml, unknown elements)
    Ignored(usize),
    /// `<none/>`
    None_,
    /// `<mprescripts/>`
    PreScripts,
}

struct Builder {
    stack: Vec<Frame>,
}

impl Builder {
    fn new() -> Self {
        Builder {
            // Start with a synthetic Math frame as root collector
            stack: vec![Frame::Math(Vec::new(), HashMap::new())],
        }
    }

    fn start_element(
        &mut self,
        name: &str,
        attrs: HashMap<String, String>,
    ) -> Result<(), MathError> {
        // If inside an ignored element, just increment depth
        if let Some(Frame::Ignored(depth)) = self.stack.last_mut() {
            *depth += 1;
            return Ok(());
        }

        match name {
            "annotation" | "annotation-xml" => {
                self.stack.push(Frame::Ignored(1));
            }
            "semantics" => {
                // Semantics is a pass-through container (SKIPPED in the
                // Python code) — we do not push a frame.  Its children
                // will be appended directly to the current parent.
            }
            "math" => {
                self.stack.push(Frame::Math(Vec::new(), attrs));
            }
            "mrow" => {
                self.stack.push(Frame::Row(Vec::new(), attrs));
            }
            "mi" => {
                self.stack
                    .push(Frame::Token(RunKind::Mi, String::new(), attrs));
            }
            "mn" => {
                self.stack
                    .push(Frame::Token(RunKind::Mn, String::new(), attrs));
            }
            "mo" => {
                self.stack
                    .push(Frame::Token(RunKind::Mo, String::new(), attrs));
            }
            "mtext" => {
                self.stack
                    .push(Frame::Token(RunKind::MText, String::new(), attrs));
            }
            "mspace" => {
                self.stack
                    .push(Frame::Token(RunKind::MSpace, String::new(), attrs));
            }
            "ms" => {
                self.stack
                    .push(Frame::Token(RunKind::Ms, String::new(), attrs));
            }
            "mfrac" => {
                self.stack.push(Frame::Frac(Vec::new(), attrs));
            }
            "msqrt" => {
                self.stack.push(Frame::Sqrt(Vec::new(), attrs));
            }
            "mroot" => {
                self.stack.push(Frame::Root(Vec::new(), attrs));
            }
            "msub" => {
                self.stack.push(Frame::Sub(Vec::new(), attrs));
            }
            "msup" => {
                self.stack.push(Frame::Sup(Vec::new(), attrs));
            }
            "msubsup" => {
                self.stack.push(Frame::SubSup(Vec::new(), attrs));
            }
            "munder" => {
                self.stack.push(Frame::Under(Vec::new(), attrs));
            }
            "mover" => {
                self.stack.push(Frame::Over(Vec::new(), attrs));
            }
            "munderover" => {
                self.stack.push(Frame::UnderOver(Vec::new(), attrs));
            }
            "mfenced" => {
                self.stack.push(Frame::Fenced(Vec::new(), attrs));
            }
            "mstyle" => {
                self.stack.push(Frame::Style(Vec::new(), attrs));
            }
            "mpadded" => {
                self.stack.push(Frame::Padded(Vec::new(), attrs));
            }
            "mphantom" => {
                self.stack.push(Frame::Phantom(Vec::new(), attrs));
            }
            "menclose" => {
                self.stack.push(Frame::Enclose(Vec::new(), attrs));
            }
            "mtable" => {
                self.stack.push(Frame::Table(Vec::new(), attrs));
            }
            "mtr" | "mlabeledtr" => {
                self.stack.push(Frame::TableRow(Vec::new(), attrs));
            }
            "mtd" => {
                self.stack.push(Frame::TableData(Vec::new(), attrs));
            }
            "maction" => {
                self.stack.push(Frame::Action(Vec::new(), attrs));
            }
            "mmultiscripts" => {
                self.stack
                    .push(Frame::MultiScripts(Vec::new(), Option::None, attrs));
            }
            "mprescripts" => {
                self.stack.push(Frame::PreScripts);
            }
            "none" => {
                self.stack.push(Frame::None_);
            }
            _ => {
                // Unknown elements: ignore for resilience
                self.stack.push(Frame::Ignored(1));
            }
        }
        Ok(())
    }

    fn end_element(&mut self, name: &str) -> Result<(), MathError> {
        // semantics is a skipped element — no frame was pushed for it
        if name == "semantics" {
            return Ok(());
        }

        // If inside an ignored block, just decrement
        if let Some(Frame::Ignored(depth)) = self.stack.last_mut() {
            *depth -= 1;
            if *depth == 0 {
                self.stack.pop();
            }
            return Ok(());
        }

        let frame = self
            .stack
            .pop()
            .ok_or_else(|| MathError::Structure(format!("unexpected closing tag: {}", name)))?;

        let node = match frame {
            Frame::Ignored(_) => return Ok(()),
            Frame::PreScripts => {
                // Signal to the parent MultiScripts
                if let Some(Frame::MultiScripts(_, ref mut prescripts, _)) = self.stack.last_mut()
                {
                    if prescripts.is_some() {
                        return Err(MathError::Structure("multiple <mprescripts/>".into()));
                    }
                    *prescripts = Some(Vec::new());
                }
                return Ok(());
            }
            Frame::None_ => OmmlNode::Empty,

            Frame::Token(kind, text, attrs) => {
                let style = match kind {
                    RunKind::Mi => {
                        let default = if text.trim().chars().count() == 1 {
                            "italic"
                        } else {
                            "normal"
                        };
                        make_run_style(&attrs, default)
                    }
                    RunKind::Mn | RunKind::Mo => make_run_style(&attrs, "normal"),
                    RunKind::MText | RunKind::Ms => {
                        let variant = attrs.get("mathvariant").map(|s| s.as_str());
                        if variant.is_none() || variant == Some("normal") {
                            RunStyle {
                                nor: true,
                                ..Default::default()
                            }
                        } else {
                            make_run_style(&attrs, "normal")
                        }
                    }
                    RunKind::MSpace => RunStyle::default(),
                };
                OmmlNode::Run {
                    kind,
                    text,
                    style,
                    attrs,
                }
            }

            Frame::Math(children, _attrs) => {
                let merged = merge_mrow_elems(children);
                OmmlNode::Math { children: merged }
            }

            Frame::Row(children, _attrs) => {
                let merged = merge_mrow_elems(children);
                if merged.len() == 1 {
                    merged.into_iter().next().unwrap()
                } else {
                    OmmlNode::Row { children: merged }
                }
            }

            Frame::Frac(children, attrs) => {
                if children.len() != 2 {
                    return Err(MathError::Structure(format!(
                        "mfrac expects 2 children, got {}",
                        children.len()
                    )));
                }
                let mut it = children.into_iter();
                OmmlNode::Frac {
                    num: Box::new(it.next().unwrap()),
                    den: Box::new(it.next().unwrap()),
                    attrs,
                }
            }

            Frame::Sqrt(children, _attrs) => {
                let merged = merge_mrow_elems(children);
                OmmlNode::Sqrt {
                    content: Box::new(row_or_single(merged)),
                }
            }

            Frame::Root(children, _attrs) => {
                if children.len() != 2 {
                    return Err(MathError::Structure(format!(
                        "mroot expects 2 children, got {}",
                        children.len()
                    )));
                }
                let mut it = children.into_iter();
                let base = it.next().unwrap();
                let degree = it.next().unwrap();
                OmmlNode::Root {
                    degree: Box::new(degree),
                    content: Box::new(base),
                }
            }

            Frame::Sub(children, _attrs) => {
                if children.len() < 2 {
                    return Err(MathError::Structure(format!(
                        "msub expects 2+ children, got {}",
                        children.len()
                    )));
                }
                let mut it = children.into_iter();
                let base = it.next().unwrap();
                let sub = it.next().unwrap();
                OmmlNode::Sub {
                    base: Box::new(base),
                    sub: Box::new(sub),
                }
            }

            Frame::Sup(children, _attrs) => {
                if children.len() < 2 {
                    return Err(MathError::Structure(format!(
                        "msup expects 2+ children, got {}",
                        children.len()
                    )));
                }
                let mut it = children.into_iter();
                let base = it.next().unwrap();
                let sup = it.next().unwrap();
                OmmlNode::Sup {
                    base: Box::new(base),
                    sup: Box::new(sup),
                }
            }

            Frame::SubSup(children, _attrs) => {
                if children.len() < 3 {
                    return Err(MathError::Structure(format!(
                        "msubsup expects 3+ children, got {}",
                        children.len()
                    )));
                }
                let mut it = children.into_iter();
                let base = it.next().unwrap();
                let sub = it.next().unwrap();
                let sup = it.next().unwrap();
                OmmlNode::SubSup {
                    base: Box::new(base),
                    sub: Box::new(sub),
                    sup: Box::new(sup),
                }
            }

            Frame::Under(children, _attrs) => {
                if children.len() < 2 {
                    return Err(MathError::Structure(format!(
                        "munder expects 2+ children, got {}",
                        children.len()
                    )));
                }
                let mut it = children.into_iter();
                let base = it.next().unwrap();
                let underscript = it.next().unwrap();
                // Check for stretchy accent -> groupChr
                if is_stretch_accent(&underscript) {
                    OmmlNode::GroupChr {
                        chr: underscript.text().to_string(),
                        pos: "bot".to_string(),
                        content: Box::new(base),
                    }
                } else {
                    OmmlNode::Under {
                        base: Box::new(base),
                        under: Box::new(underscript),
                    }
                }
            }

            Frame::Over(children, _attrs) => {
                if children.len() < 2 {
                    return Err(MathError::Structure(format!(
                        "mover expects 2+ children, got {}",
                        children.len()
                    )));
                }
                let mut it = children.into_iter();
                let base = it.next().unwrap();
                let overscript = it.next().unwrap();
                if is_stretch_accent(&overscript) {
                    OmmlNode::GroupChr {
                        chr: overscript.text().to_string(),
                        pos: "top".to_string(),
                        content: Box::new(base),
                    }
                } else {
                    OmmlNode::Over {
                        base: Box::new(base),
                        over: Box::new(overscript),
                    }
                }
            }

            Frame::UnderOver(children, _attrs) => {
                if children.len() < 3 {
                    return Err(MathError::Structure(format!(
                        "munderover expects 3+ children, got {}",
                        children.len()
                    )));
                }
                let mut it = children.into_iter();
                let base = it.next().unwrap();
                let under = it.next().unwrap();
                let over = it.next().unwrap();
                OmmlNode::UnderOver {
                    base: Box::new(base),
                    under: Box::new(under),
                    over: Box::new(over),
                }
            }

            Frame::Fenced(children, attrs) => {
                let open = attrs
                    .get("open")
                    .cloned()
                    .unwrap_or_else(|| "(".to_string());
                let close = attrs
                    .get("close")
                    .cloned()
                    .unwrap_or_else(|| ")".to_string());
                let seps = attrs
                    .get("separators")
                    .cloned()
                    .unwrap_or_else(|| ",".to_string());
                OmmlNode::Delimited {
                    open,
                    close,
                    separators: seps,
                    children,
                }
            }

            Frame::Style(children, _attrs) => {
                let merged = merge_mrow_elems(children);
                OmmlNode::Style {
                    content: Box::new(row_or_single(merged)),
                }
            }

            Frame::Padded(children, _attrs) => {
                let merged = merge_mrow_elems(children);
                OmmlNode::Padded {
                    content: Box::new(row_or_single(merged)),
                }
            }

            Frame::Phantom(children, _attrs) => {
                let merged = merge_mrow_elems(children);
                OmmlNode::Phantom {
                    content: Box::new(row_or_single(merged)),
                }
            }

            Frame::Enclose(children, _attrs) => {
                let merged = merge_mrow_elems(children);
                OmmlNode::Enclose {
                    content: Box::new(row_or_single(merged)),
                }
            }

            Frame::Table(rows, _attrs) => OmmlNode::Matrix { rows },

            Frame::TableRow(children, _attrs) => {
                // Push onto parent Table
                if let Some(Frame::Table(ref mut rows, _)) = self.stack.last_mut() {
                    rows.push(children);
                    return Ok(());
                }
                return Err(MathError::Structure("mtr outside mtable".into()));
            }

            Frame::TableData(children, _attrs) => {
                let merged = merge_mrow_elems(children);
                let node = row_or_single(merged);
                // Push onto parent TableRow
                if let Some(Frame::TableRow(ref mut cells, _)) = self.stack.last_mut() {
                    cells.push(node);
                    return Ok(());
                }
                return Err(MathError::Structure("mtd outside mtr".into()));
            }

            Frame::Action(children, attrs) => {
                if children.is_empty() {
                    return Err(MathError::Structure("maction with no children".into()));
                }
                let action_type = attrs
                    .get("actiontype")
                    .map(|s| s.as_str())
                    .unwrap_or("toggle");
                let idx = if action_type == "toggle" {
                    let sel: usize = attrs
                        .get("selection")
                        .and_then(|s| s.parse().ok())
                        .unwrap_or(1);
                    sel.saturating_sub(1).min(children.len() - 1)
                } else {
                    0
                };
                let selected = children.into_iter().nth(idx).unwrap();
                OmmlNode::Action {
                    selected: Box::new(selected),
                }
            }

            Frame::MultiScripts(children, prescripts, _attrs) => {
                if children.is_empty() {
                    return Err(MathError::Structure(
                        "mmultiscripts with no children".into(),
                    ));
                }
                let mut it = children.into_iter();
                let base = it.next().unwrap();
                let post: Vec<OmmlNode> = it.collect();
                let mut post_pairs = Vec::new();
                let mut pi = post.into_iter();
                while let Some(sub) = pi.next() {
                    let sup = pi.next().unwrap_or(OmmlNode::Empty);
                    post_pairs.push((sub, sup));
                }
                let mut pre_pairs = Vec::new();
                if let Some(pre) = prescripts {
                    let mut pi = pre.into_iter();
                    while let Some(sub) = pi.next() {
                        let sup = pi.next().unwrap_or(OmmlNode::Empty);
                        pre_pairs.push((sub, sup));
                    }
                }
                OmmlNode::MultiScripts {
                    base: Box::new(base),
                    post_pairs,
                    pre_pairs,
                }
            }
        };

        // Push the finished node onto the parent frame
        self.push_child(node)?;
        Ok(())
    }

    fn characters(&mut self, content: &str) {
        match self.stack.last_mut() {
            Some(Frame::Token(_, ref mut text, _)) => {
                text.push_str(content);
            }
            _ => {
                // Non-whitespace text outside a token element is ignored
                // (matches Python behaviour which just prints it).
            }
        }
    }

    fn push_child(&mut self, node: OmmlNode) -> Result<(), MathError> {
        match self.stack.last_mut() {
            Some(Frame::Math(ref mut v, _))
            | Some(Frame::Row(ref mut v, _))
            | Some(Frame::Frac(ref mut v, _))
            | Some(Frame::Sqrt(ref mut v, _))
            | Some(Frame::Root(ref mut v, _))
            | Some(Frame::Sub(ref mut v, _))
            | Some(Frame::Sup(ref mut v, _))
            | Some(Frame::SubSup(ref mut v, _))
            | Some(Frame::Under(ref mut v, _))
            | Some(Frame::Over(ref mut v, _))
            | Some(Frame::UnderOver(ref mut v, _))
            | Some(Frame::Fenced(ref mut v, _))
            | Some(Frame::Style(ref mut v, _))
            | Some(Frame::Padded(ref mut v, _))
            | Some(Frame::Phantom(ref mut v, _))
            | Some(Frame::Enclose(ref mut v, _))
            | Some(Frame::TableRow(ref mut v, _))
            | Some(Frame::TableData(ref mut v, _))
            | Some(Frame::Action(ref mut v, _)) => {
                v.push(node);
            }
            Some(Frame::MultiScripts(ref mut children, ref mut prescripts, _)) => {
                if let Some(ref mut pre) = prescripts {
                    pre.push(node);
                } else {
                    children.push(node);
                }
            }
            // semantics is skipped (no frame pushed); this arm is unreachable
            Some(Frame::Table(_, _)) => {
                // Direct children of table should be mtr; ignore others
            }
            Some(Frame::Ignored(_)) | Some(Frame::Token(..)) => {}
            Some(Frame::None_) | Some(Frame::PreScripts) => {}
            None => {
                return Err(MathError::Structure("no parent frame for child".into()));
            }
        }
        Ok(())
    }

    fn result(mut self) -> Result<OmmlNode, MathError> {
        if self.stack.len() != 1 {
            return Err(MathError::Structure(format!(
                "unclosed elements: {} frames remain on stack",
                self.stack.len()
            )));
        }
        match self.stack.pop().unwrap() {
            Frame::Math(children, _) => {
                let merged = merge_mrow_elems(children);
                Ok(OmmlNode::Math { children: merged })
            }
            _ => Err(MathError::Structure("expected <math> as root".into())),
        }
    }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/// Common MathML named entities and their Unicode code points.
const MATH_ENTITIES: &[(&str, char)] = &[
    ("pi", '\u{03C0}'),
    ("ExponentialE", '\u{2147}'),
    ("ee", '\u{2147}'),
    ("ImaginaryI", '\u{2148}'),
    ("ii", '\u{2148}'),
    ("gamma", '\u{03B3}'),
    ("infin", '\u{221E}'),
    ("infty", '\u{221E}'),
];

/// Replace named entity references (`&pi;`, `&infin;`, etc.) with their
/// Unicode characters.  `quick-xml` does not resolve custom DTD entities,
/// so we do it as a pre-processing step.
fn resolve_entities(input: &str) -> String {
    let mut result = input.to_string();
    for &(name, ch) in MATH_ENTITIES {
        let entity = format!("&{};", name);
        if result.contains(&entity) {
            result = result.replace(&entity, &ch.to_string());
        }
    }
    result
}

/// Convert a MathML XML string to an Office MathML (OMML) XML string.
///
/// The input should be a `<math>` element (with or without namespace).
/// Common MathML named entities (`&pi;`, `&infin;`, etc.) are supported
/// automatically.
///
/// # Example
///
/// ```ignore
/// let omml = deckmint_math::mathml_to_omml::convert(
///     "<math><mi>x</mi></math>"
/// ).unwrap();
/// assert!(omml.contains("<m:oMath>"));
/// ```
pub fn convert(mathml: &str) -> Result<String, MathError> {
    // Resolve named entities before parsing (quick-xml doesn't handle
    // custom DTD entities).
    let input = resolve_entities(mathml);

    let mut reader = Reader::from_str(&input);
    reader.config_mut().trim_text(false);

    let mut builder = Builder::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = local_name(e.name().as_ref());
                let attrs = collect_attrs(e)?;
                builder.start_element(&name, attrs)?;
            }
            Ok(Event::Empty(ref e)) => {
                let name = local_name(e.name().as_ref());
                let attrs = collect_attrs(e)?;
                builder.start_element(&name, attrs)?;
                builder.end_element(&name)?;
            }
            Ok(Event::End(ref e)) => {
                let name = local_name(e.name().as_ref());
                builder.end_element(&name)?;
            }
            Ok(Event::Text(ref e)) => {
                let text = e
                    .unescape()
                    .map_err(|err| MathError::Xml(err.to_string()))?;
                builder.characters(&text);
            }
            Ok(Event::Eof) => break,
            Ok(Event::Decl(_))
            | Ok(Event::PI(_))
            | Ok(Event::Comment(_))
            | Ok(Event::DocType(_)) => {
                // skip
            }
            Ok(Event::CData(ref e)) => {
                let text = String::from_utf8_lossy(e.as_ref());
                builder.characters(&text);
            }
            Err(e) => return Err(MathError::Xml(format!("{}", e))),
        }
        buf.clear();
    }

    let root = builder.result()?;
    Ok(root.to_omml())
}

/// Extract the local name from a potentially namespaced tag name.
fn local_name(raw: &[u8]) -> String {
    let full = std::str::from_utf8(raw).unwrap_or("");
    // Strip namespace prefix (e.g. "m:math" -> "math", "mml:mrow" -> "mrow")
    match full.rfind(':') {
        Some(pos) => full[pos + 1..].to_string(),
        None => full.to_string(),
    }
}

/// Collect attributes from a `BytesStart` (opening or empty-element tag).
fn collect_attrs(e: &quick_xml::events::BytesStart<'_>) -> Result<HashMap<String, String>, MathError> {
    let mut map = HashMap::new();
    for attr in e.attributes() {
        let attr = attr.map_err(|err| MathError::Xml(format!("attr error: {}", err)))?;
        let key = local_name(attr.key.as_ref());
        let val = attr
            .unescape_value()
            .map_err(|err| MathError::Xml(err.to_string()))?;
        map.insert(key, val.to_string());
    }
    Ok(map)
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_identifier() {
        let omml = convert("<math><mi>x</mi></math>").unwrap();
        assert!(omml.contains("<m:oMath>"), "missing oMath wrapper");
        assert!(omml.contains("<m:t>x</m:t>"), "missing text");
    }

    #[test]
    fn test_fraction() {
        let omml = convert("<math><mfrac><mn>1</mn><mn>2</mn></mfrac></math>").unwrap();
        assert!(omml.contains("<m:f>"), "missing f element");
        assert!(omml.contains("<m:num>"), "missing num");
        assert!(omml.contains("<m:den>"), "missing den");
    }

    #[test]
    fn test_sqrt() {
        let omml = convert("<math><msqrt><mi>x</mi></msqrt></math>").unwrap();
        assert!(omml.contains("<m:rad>"), "missing rad element");
    }

    #[test]
    fn test_subscript() {
        let omml = convert("<math><msub><mi>x</mi><mn>2</mn></msub></math>").unwrap();
        assert!(omml.contains("<m:sSub>"), "missing sSub");
        assert!(omml.contains("<m:sub>"), "missing sub");
    }

    #[test]
    fn test_superscript() {
        let omml = convert("<math><msup><mi>x</mi><mn>2</mn></msup></math>").unwrap();
        assert!(omml.contains("<m:sSup>"), "missing sSup");
        assert!(omml.contains("<m:sup>"), "missing sup");
    }

    #[test]
    fn test_subsup() {
        let omml =
            convert("<math><msubsup><mi>x</mi><mn>1</mn><mn>2</mn></msubsup></math>").unwrap();
        assert!(omml.contains("<m:sSubSup>"), "missing sSubSup");
    }

    #[test]
    fn test_fenced() {
        let omml = convert("<math><mfenced><mi>x</mi></mfenced></math>").unwrap();
        assert!(omml.contains("<m:d>"), "missing d element");
        assert!(omml.contains("<m:begChr"), "missing begChr");
        assert!(omml.contains("<m:endChr"), "missing endChr");
    }

    #[test]
    fn test_invisible_chars_stripped() {
        let omml = convert("<math><mo>\u{2062}</mo></math>").unwrap();
        assert!(omml.contains("<m:t></m:t>"), "invisible char not stripped");
    }

    #[test]
    fn test_entity_pi() {
        let omml = convert("<math><mi>&pi;</mi></math>").unwrap();
        assert!(
            omml.contains("\u{03C0}") || omml.contains("π"),
            "pi entity not resolved"
        );
    }

    #[test]
    fn test_mrow_passthrough() {
        let omml =
            convert("<math><mrow><mn>1</mn><mo>+</mo><mn>2</mn></mrow></math>").unwrap();
        assert!(omml.contains("<m:oMath>"));
    }

    #[test]
    fn test_matrix() {
        let omml = convert(
            "<math><mtable><mtr><mtd><mn>1</mn></mtd><mtd><mn>2</mn></mtd></mtr></mtable></math>",
        )
        .unwrap();
        assert!(omml.contains("<m:m>"), "missing m element");
        assert!(omml.contains("<m:mr>"), "missing mr element");
    }

    #[test]
    fn test_root() {
        let omml = convert("<math><mroot><mi>x</mi><mn>3</mn></mroot></math>").unwrap();
        assert!(omml.contains("<m:rad>"), "missing rad");
        assert!(omml.contains("<m:deg>"), "missing deg");
    }

    #[test]
    fn test_phantom() {
        let omml = convert("<math><mphantom><mi>x</mi></mphantom></math>").unwrap();
        assert!(omml.contains("<m:phant>"), "missing phant");
        assert!(omml.contains("<m:show m:val=\"0\"/>"), "missing show=0");
    }

    #[test]
    fn test_xml_escaping() {
        let omml = convert("<math><mo>&lt;</mo></math>").unwrap();
        assert!(omml.contains("&lt;"), "< not escaped");
    }

    #[test]
    fn test_under() {
        let omml = convert("<math><munder><mi>x</mi><mo>_</mo></munder></math>").unwrap();
        assert!(omml.contains("<m:limLow>"), "missing limLow");
    }

    #[test]
    fn test_over() {
        let omml = convert("<math><mover><mi>x</mi><mo>^</mo></mover></math>").unwrap();
        assert!(omml.contains("<m:limUpp>"), "missing limUpp");
    }

    #[test]
    fn test_underover() {
        let omml = convert(
            "<math><munderover><mi>x</mi><mn>0</mn><mn>1</mn></munderover></math>",
        )
        .unwrap();
        assert!(omml.contains("<m:limLow>"), "missing limLow");
        assert!(omml.contains("<m:limUpp>"), "missing limUpp");
    }

    #[test]
    fn test_mstyle_passthrough() {
        let omml = convert("<math><mstyle><mi>x</mi></mstyle></math>").unwrap();
        assert!(omml.contains("<m:t>x</m:t>"));
    }

    #[test]
    fn test_semantics_passthrough() {
        let omml = convert(
            "<math><semantics><mi>x</mi><annotation encoding=\"TeX\">x</annotation></semantics></math>",
        )
        .unwrap();
        assert!(omml.contains("<m:t>x</m:t>"));
        assert!(!omml.contains("TeX"));
    }

    #[test]
    fn test_ms_quotes() {
        let omml = convert("<math><ms>hello</ms></math>").unwrap();
        assert!(
            omml.contains("&quot;hello&quot;"),
            "ms quotes not rendered"
        );
    }

    #[test]
    fn test_enclose() {
        let omml = convert("<math><menclose><mi>x</mi></menclose></math>").unwrap();
        assert!(omml.contains("<m:borderBox>"), "missing borderBox");
    }

    #[test]
    fn test_nobar_frac() {
        let omml = convert(
            "<math><mfrac linethickness=\"0px\"><mn>1</mn><mn>2</mn></mfrac></math>",
        )
        .unwrap();
        assert!(omml.contains("noBar"), "missing noBar for zero thickness");
    }
}
