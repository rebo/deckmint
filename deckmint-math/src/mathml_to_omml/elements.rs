// elements.rs — OMML element types ported from Python mathml2omml
//
// Each variant of `OmmlNode` corresponds to a MathML element class in the
// original Python code.  The `to_omml` method on `OmmlNode` emits the
// Office MathML (OMML) XML fragment for that element.

use std::collections::HashMap;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/// Escape XML special characters (`&`, `<`, `>`, `"`, `'`).
pub fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

/// Strip Unicode invisible characters (U+2061 .. U+2064).
pub fn strip_invisible(s: &str) -> String {
    s.chars()
        .filter(|&c| !('\u{2061}'..='\u{2064}').contains(&c))
        .collect()
}

/// Escape text for OMML output: strip invisible chars then XML-escape.
pub fn escape_text(s: &str) -> String {
    xml_escape(&strip_invisible(s))
}

// ---------------------------------------------------------------------------
// Run style
// ---------------------------------------------------------------------------

/// Style properties attached to a math run (`<m:rPr>`).
#[derive(Debug, Clone, Default)]
pub struct RunStyle {
    /// e.g. `"script"`, `"fraktur"`, `"double-struck"`, etc.
    pub script: Option<String>,
    /// One of `"bi"`, `"b"`, `"i"`, `"p"` (plain/normal).
    pub style: Option<String>,
    /// If true, emit `<m:nor/>` instead of scr/sty.
    pub nor: bool,
}

impl RunStyle {
    /// Build the inner XML for `<m:rPr>`.
    pub fn to_omml_inner(&self) -> String {
        if self.nor {
            return "<m:nor/>".to_string();
        }
        let mut out = String::new();
        if let Some(ref scr) = self.script {
            out.push_str(&format!("<m:scr m:val=\"{}\"/>", scr));
        }
        if let Some(ref sty) = self.style {
            out.push_str(&format!("<m:sty m:val=\"{}\"/>", sty));
        }
        out
    }
}

// ---------------------------------------------------------------------------
// mathvariant helpers
// ---------------------------------------------------------------------------

static SCRIPT_LIST: &[&str] = &[
    "script", "fraktur", "sans-serif", "monospace", "double-struck", "roman",
];

/// Map a `mathvariant` attribute value to a `RunStyle`.
pub fn make_run_style(attrs: &HashMap<String, String>, default_style: &str) -> RunStyle {
    let style_map: &[(&str, &str)] = &[
        ("bold-italic", "bi"),
        ("bold", "b"),
        ("italic", "i"),
        ("normal", "p"),
    ];

    let variant = match attrs.get("mathvariant") {
        Some(v) => v.clone(),
        None => {
            return RunStyle {
                script: None,
                style: Some(
                    style_map
                        .iter()
                        .find(|(k, _)| *k == default_style)
                        .map(|(_, v)| v.to_string())
                        .unwrap_or_else(|| "p".to_string()),
                ),
                nor: false,
            };
        }
    };

    let mut rs = RunStyle::default();
    for scr in SCRIPT_LIST {
        if variant.contains(scr) {
            rs.script = Some(scr.to_string());
            break;
        }
    }
    let mut found_style = false;
    for &(style, val) in style_map {
        if variant.contains(style) {
            rs.style = Some(val.to_string());
            found_style = true;
            break;
        }
    }
    if !found_style {
        rs.style = Some(
            style_map
                .iter()
                .find(|(k, _)| *k == default_style)
                .map(|(_, v)| v.to_string())
                .unwrap_or_else(|| "p".to_string()),
        );
    }
    rs
}

// ---------------------------------------------------------------------------
// RunKind — distinguishes MI / MN / MO / MText / MSpace / MS
// ---------------------------------------------------------------------------

/// The kind of math run, corresponding to MathML token elements.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RunKind {
    /// `<mi>` — identifier
    Mi,
    /// `<mn>` — number
    Mn,
    /// `<mo>` — operator / fence / separator
    Mo,
    /// `<mtext>` — text
    MText,
    /// `<mspace>` — space
    MSpace,
    /// `<ms>` — string literal
    Ms,
}

// ---------------------------------------------------------------------------
// OmmlNode
// ---------------------------------------------------------------------------

/// A node in the OMML element tree.
///
/// Each variant maps to a MathML element and knows how to serialise itself
/// to Office MathML XML via [`OmmlNode::to_omml`].
#[derive(Debug, Clone)]
pub enum OmmlNode {
    /// A math run (token element): MI, MN, MO, MText, MSpace, MS.
    Run {
        kind: RunKind,
        text: String,
        style: RunStyle,
        attrs: HashMap<String, String>,
    },

    /// `<mrow>` — horizontal group of sub-expressions.
    Row {
        children: Vec<OmmlNode>,
    },

    /// `<mfrac>` — fraction.
    Frac {
        num: Box<OmmlNode>,
        den: Box<OmmlNode>,
        attrs: HashMap<String, String>,
    },

    /// `<msqrt>` — square root (inferred mrow content).
    Sqrt {
        content: Box<OmmlNode>,
    },

    /// `<mroot>` — n-th root.
    Root {
        degree: Box<OmmlNode>,
        content: Box<OmmlNode>,
    },

    /// `<msub>` — subscript.
    Sub {
        base: Box<OmmlNode>,
        sub: Box<OmmlNode>,
    },

    /// `<msup>` — superscript.
    Sup {
        base: Box<OmmlNode>,
        sup: Box<OmmlNode>,
    },

    /// `<msubsup>` — subscript-superscript pair.
    SubSup {
        base: Box<OmmlNode>,
        sub: Box<OmmlNode>,
        sup: Box<OmmlNode>,
    },

    /// `<munder>` — underscript.
    Under {
        base: Box<OmmlNode>,
        under: Box<OmmlNode>,
    },

    /// `<mover>` — overscript.
    Over {
        base: Box<OmmlNode>,
        over: Box<OmmlNode>,
    },

    /// `<munderover>` — underscript-overscript pair.
    UnderOver {
        base: Box<OmmlNode>,
        under: Box<OmmlNode>,
        over: Box<OmmlNode>,
    },

    /// `<mfenced>` or synthesised delimited expression.
    Delimited {
        open: String,
        close: String,
        separators: String,
        children: Vec<OmmlNode>,
    },

    /// `<menclose>` — border box.
    Enclose {
        content: Box<OmmlNode>,
    },

    /// `<mphantom>` — invisible content.
    Phantom {
        content: Box<OmmlNode>,
    },

    /// `<mstyle>` — style wrapper (pass-through).
    Style {
        content: Box<OmmlNode>,
    },

    /// `<mpadded>` — padded wrapper (pass-through).
    Padded {
        content: Box<OmmlNode>,
    },

    /// `<mtable>` — matrix / table.
    Matrix {
        rows: Vec<Vec<OmmlNode>>,
    },

    /// `<maction>` — action binding (renders selected child).
    Action {
        selected: Box<OmmlNode>,
    },

    /// `<mmultiscripts>` — prescripts and tensor indices.
    MultiScripts {
        base: Box<OmmlNode>,
        post_pairs: Vec<(OmmlNode, OmmlNode)>,
        pre_pairs: Vec<(OmmlNode, OmmlNode)>,
    },

    /// N-ary operator (promoted from Sub/Sup/SubSup/Under/Over/UnderOver
    /// when the base is a large operator).
    Nary {
        chr: String,
        sub: Box<OmmlNode>,
        sup: Box<OmmlNode>,
        content: Box<OmmlNode>,
        lim_loc: String,
    },

    /// `<m:groupChr>` — stretchy accent above or below.
    GroupChr {
        chr: String,
        pos: String,
        content: Box<OmmlNode>,
    },

    /// Top-level `<math>` element producing `<m:oMath>`.
    Math {
        children: Vec<OmmlNode>,
    },

    /// An empty element (used as a placeholder, e.g. empty sub/sup in N-ary).
    Empty,
}

impl OmmlNode {
    // -- convenience constructors -------------------------------------------

    /// Create an empty `Row`.
    pub fn empty_row() -> Self {
        OmmlNode::Row {
            children: Vec::new(),
        }
    }

    /// Create a `Run` of kind `Mo` with the given text.
    pub fn mo(text: &str, attrs: HashMap<String, String>) -> Self {
        let style = make_run_style(&attrs, "normal");
        OmmlNode::Run {
            kind: RunKind::Mo,
            text: text.to_string(),
            style,
            attrs,
        }
    }

    // -- text access --------------------------------------------------------

    /// Return the raw text of a `Run` node, or empty string otherwise.
    pub fn text(&self) -> &str {
        match self {
            OmmlNode::Run { text, .. } => text.as_str(),
            _ => "",
        }
    }

    /// Return the escaped text (invisible chars stripped, XML-escaped).
    pub fn escaped_text(&self) -> String {
        escape_text(self.text())
    }

    /// For an `Ms` run, return text with surrounding quotes.
    pub fn ms_text(&self) -> String {
        if let OmmlNode::Run {
            kind: RunKind::Ms,
            text,
            attrs,
            ..
        } = self
        {
            let lquote = attrs.get("lquote").map(|s| s.as_str()).unwrap_or("\"");
            let rquote = attrs.get("rquote").map(|s| s.as_str()).unwrap_or("\"");
            format!("{}{}{}", lquote, text, rquote)
        } else {
            self.text().to_string()
        }
    }

    // -- predicates ---------------------------------------------------------

    /// Whether this node is "space-like" per MathML spec.
    pub fn is_space_like(&self) -> bool {
        match self {
            OmmlNode::Run { kind, .. } => matches!(kind, RunKind::MText | RunKind::MSpace),
            OmmlNode::Row { children } => children.iter().all(|c| c.is_space_like()),
            OmmlNode::Style { content } | OmmlNode::Padded { content } => content.is_space_like(),
            OmmlNode::Phantom { content } => content.is_space_like(),
            OmmlNode::Empty => true,
            _ => false,
        }
    }

    /// Return a reference to the core embellished operator, if any.
    pub fn embellished_op(&self) -> Option<&OmmlNode> {
        match self {
            OmmlNode::Run {
                kind: RunKind::Mo, ..
            } => Some(self),
            OmmlNode::Row { children } => {
                let mut first_emb: Option<&OmmlNode> = None;
                for child in children {
                    if child.is_space_like() {
                        continue;
                    }
                    let emb = child.embellished_op();
                    if emb.is_none() || first_emb.is_some() {
                        return None;
                    }
                    first_emb = emb;
                }
                first_emb
            }
            OmmlNode::Sub { .. }
            | OmmlNode::Sup { .. }
            | OmmlNode::SubSup { .. }
            | OmmlNode::Under { .. }
            | OmmlNode::Over { .. }
            | OmmlNode::UnderOver { .. }
            | OmmlNode::Frac { .. }
            | OmmlNode::MultiScripts { .. }
            | OmmlNode::Phantom { .. } => Some(self),
            OmmlNode::Style { content } | OmmlNode::Padded { content } => content.embellished_op(),
            _ => None,
        }
    }

    // -- OMML serialisation -------------------------------------------------

    /// Serialise this node to an OMML XML string.
    pub fn to_omml(&self) -> String {
        match self {
            OmmlNode::Run {
                kind,
                text,
                style,
                attrs,
            } => {
                let display_text = if *kind == RunKind::Ms {
                    let lquote = attrs.get("lquote").map(|s| s.as_str()).unwrap_or("\"");
                    let rquote = attrs.get("rquote").map(|s| s.as_str()).unwrap_or("\"");
                    escape_text(&format!("{}{}{}", lquote, text, rquote))
                } else {
                    escape_text(text)
                };
                let props_inner = style.to_omml_inner();
                let run_prop = if props_inner.is_empty() {
                    String::new()
                } else {
                    format!("<m:rPr>{}</m:rPr>", props_inner)
                };
                format!("<m:r>{}<m:t>{}</m:t></m:r>", run_prop, display_text)
            }

            OmmlNode::Row { children } => {
                if children.is_empty() {
                    return String::new();
                }
                let inner: String = children.iter().map(|c| to_box(c)).collect();
                format!("<m:box><m:e>{}</m:e></m:box>", inner)
            }

            OmmlNode::Frac { num, den, attrs } => {
                let props = frac_props(attrs);
                let num_s = to_math_arg(num);
                let den_s = to_math_arg(den);
                format!(
                    "<m:f>{}<m:num>{}</m:num><m:den>{}</m:den></m:f>",
                    props, num_s, den_s
                )
            }

            OmmlNode::Sqrt { content } => {
                let elem = to_element(content);
                format!("<m:rad>{}</m:rad>", elem)
            }

            OmmlNode::Root { degree, content } => {
                let deg = to_math_arg(degree);
                let elem = to_element(content);
                format!("<m:rad><m:deg>{}</m:deg>{}</m:rad>", deg, elem)
            }

            OmmlNode::Sub { base, sub } => {
                let base_s = to_element(base);
                let sub_s = to_math_arg(sub);
                format!("<m:sSub>{}<m:sub>{}</m:sub></m:sSub>", base_s, sub_s)
            }

            OmmlNode::Sup { base, sup } => {
                let base_s = to_element(base);
                let sup_s = to_math_arg(sup);
                format!("<m:sSup>{}<m:sup>{}</m:sup></m:sSup>", base_s, sup_s)
            }

            OmmlNode::SubSup { base, sub, sup } => {
                let base_s = to_element(base);
                let sub_s = to_math_arg(sub);
                let sup_s = to_math_arg(sup);
                format!(
                    "<m:sSubSup>{}<m:sub>{}</m:sub><m:sup>{}</m:sup></m:sSubSup>",
                    base_s, sub_s, sup_s
                )
            }

            OmmlNode::Under { base, under } => {
                let base_s = to_element(base);
                let under_s = to_math_arg(under);
                format!(
                    "<m:limLow>{}<m:lim>{}</m:lim></m:limLow>",
                    base_s, under_s
                )
            }

            OmmlNode::Over { base, over } => {
                let base_s = to_element(base);
                let over_s = to_math_arg(over);
                format!(
                    "<m:limUpp>{}<m:lim>{}</m:lim></m:limUpp>",
                    base_s, over_s
                )
            }

            OmmlNode::UnderOver {
                base,
                under,
                over,
            } => {
                let base_s = to_element(base);
                let over_s = to_math_arg(over);
                let under_s = to_math_arg(under);
                format!(
                    "<m:limLow><m:e><m:limUpp>{}<m:lim>{}</m:lim></m:limUpp></m:e><m:lim>{}</m:lim></m:limLow>",
                    base_s, over_s, under_s
                )
            }

            OmmlNode::Delimited {
                open,
                close,
                separators,
                children,
            } => {
                let beg = xml_escape(open);
                let end = xml_escape(close);
                let seps: Vec<char> =
                    separators.chars().filter(|c| !c.is_whitespace()).collect();

                let elem_str = if children.len() > 1 && !seps.is_empty() {
                    // Interleave children with separator runs
                    let mut parts = Vec::new();
                    for (idx, child) in children.iter().enumerate() {
                        if idx > 0 {
                            let sep_ch = if idx - 1 < seps.len() {
                                seps[idx - 1]
                            } else {
                                *seps.last().unwrap()
                            };
                            parts.push(OmmlNode::Run {
                                kind: RunKind::Mo,
                                text: sep_ch.to_string(),
                                style: RunStyle::default(),
                                attrs: HashMap::new(),
                            });
                        }
                        parts.push(child.clone());
                    }
                    let combined = OmmlNode::Row { children: parts };
                    to_element_from_children(&[combined])
                } else {
                    to_element_from_children(children)
                };

                format!(
                    "<m:d><m:dPr><m:begChr m:val=\"{}\"/><m:endChr m:val=\"{}\"/></m:dPr>{}</m:d>",
                    beg, end, elem_str
                )
            }

            OmmlNode::Enclose { content } => {
                let elem = to_element(content);
                format!("<m:borderBox>{}</m:borderBox>", elem)
            }

            OmmlNode::Phantom { content } => {
                let elem = to_element(content);
                format!(
                    "<m:phant><m:phantPr><m:show m:val=\"0\"/></m:phantPr>{}</m:phant>",
                    elem
                )
            }

            OmmlNode::Style { content } | OmmlNode::Padded { content } => {
                // Pass-through: just render the inner content
                content.to_omml()
            }

            OmmlNode::Matrix { rows } => {
                let max_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
                let rows_str: String = rows
                    .iter()
                    .map(|row| {
                        let mut cells: Vec<String> = row.iter().map(|c| to_element(c)).collect();
                        // Pad short rows
                        while cells.len() < max_cols {
                            cells.push("<m:e></m:e>".to_string());
                        }
                        format!("<m:mr>{}</m:mr>", cells.join(""))
                    })
                    .collect();
                format!("<m:m>{}</m:m>", rows_str)
            }

            OmmlNode::Action { selected } => selected.to_omml(),

            OmmlNode::MultiScripts {
                base,
                post_pairs,
                pre_pairs,
            } => {
                // Build from inside out: base -> post-scripts -> pre-scripts
                let mut result = base.to_omml();
                for (postsub, postsup) in post_pairs {
                    let sub_s = to_math_arg(postsub);
                    let sup_s = to_math_arg(postsup);
                    result = format!(
                        "<m:sSubSup>{}<m:sub>{}</m:sub><m:sup>{}</m:sup></m:sSubSup>",
                        wrap_element(&result),
                        sub_s,
                        sup_s
                    );
                }
                if pre_pairs.is_empty() {
                    return format!("<m:sPre>{}</m:sPre>", wrap_element(&result));
                }
                for (presub, presup) in pre_pairs.iter().rev() {
                    let sub_s = to_math_arg(presub);
                    let sup_s = to_math_arg(presup);
                    result = format!(
                        "<m:sPre><m:sub>{}</m:sub><m:sup>{}</m:sup>{}</m:sPre>",
                        sub_s,
                        sup_s,
                        wrap_element(&result)
                    );
                }
                result
            }

            OmmlNode::Nary {
                chr,
                sub,
                sup,
                content,
                lim_loc,
            } => {
                let sub_s = to_math_arg(sub);
                let sup_s = to_math_arg(sup);
                let elem = to_element(content);
                format!(
                    "<m:nary><m:naryPr><m:chr m:val=\"{}\"/><m:limLoc m:val=\"{}\"/></m:naryPr><m:sub>{}</m:sub><m:sup>{}</m:sup>{}</m:nary>",
                    escape_text(chr), lim_loc, sub_s, sup_s, elem
                )
            }

            OmmlNode::GroupChr { chr, pos, content } => {
                let elem = to_element(content);
                format!(
                    "<m:groupChr><m:groupChrPr><m:chr m:val=\"{}\"/><m:pos m:val=\"{}\"/></m:groupChr>{}</m:groupChr>",
                    escape_text(chr), pos, elem
                )
            }

            OmmlNode::Math { children } => {
                let inner: String = children.iter().map(|c| to_box(c)).collect();
                format!("<m:oMath>{}</m:oMath>", inner)
            }

            OmmlNode::Empty => String::new(),
        }
    }
}

// ---------------------------------------------------------------------------
// Helper functions mirroring Python's to_box / to_element / to_math_arg
// ---------------------------------------------------------------------------

/// Wrap a node in `<m:box>` if it is an `Element`-like (Row), otherwise
/// just serialise it.  Mirrors Python `to_box`.
pub fn to_box(node: &OmmlNode) -> String {
    match node {
        OmmlNode::Row { children } => {
            let inner: String = children.iter().map(|c| to_box(c)).collect();
            format!("<m:box><m:e>{}</m:e></m:box>", inner)
        }
        _ => node.to_omml(),
    }
}

/// Wrap a node inside `<m:e>...</m:e>`.  Mirrors Python `to_element`.
pub fn to_element(node: &OmmlNode) -> String {
    match node {
        OmmlNode::Row { children } => {
            let inner: String = children.iter().map(|c| to_box(c)).collect();
            format!("<m:e>{}</m:e>", inner)
        }
        _ => format!("<m:e>{}</m:e>", to_box(node)),
    }
}

/// Produce `<m:e>` from a slice of children.
fn to_element_from_children(children: &[OmmlNode]) -> String {
    let inner: String = children.iter().map(|c| to_box(c)).collect();
    format!("<m:e>{}</m:e>", inner)
}

/// Wrap raw OMML string in `<m:e>`.
fn wrap_element(omml: &str) -> String {
    format!("<m:e>{}</m:e>", omml)
}

/// Extract the "math arg" content of a node: if it is a Row, emit each
/// child boxed; otherwise emit the node directly.  Mirrors Python
/// `to_math_arg`.
pub fn to_math_arg(node: &OmmlNode) -> String {
    match node {
        OmmlNode::Row { children } => {
            children.iter().map(|c| to_box(c)).collect()
        }
        _ => node.to_omml(),
    }
}

// ---------------------------------------------------------------------------
// Fraction properties helper
// ---------------------------------------------------------------------------

fn frac_props(attrs: &HashMap<String, String>) -> String {
    if let Some(thickness) = attrs.get("linethickness") {
        let t = thickness.trim();
        // "0", "0px", "0em", etc. → noBar
        if t.starts_with('0') && !t.chars().nth(1).map_or(false, |c| c.is_ascii_digit() || c == '.') {
            return "<m:fPr><m:type m:val=\"noBar\"/></m:fPr>".to_string();
        }
    }
    if attrs.get("bevelled").map(|v| v.as_str()) == Some("true") {
        return "<m:fPr><m:type m:val=\"skw\"/></m:fPr>".to_string();
    }
    String::new()
}
