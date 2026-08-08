// postprocess.rs — PowerPoint-specific post-processing for OMML equations
//
// The raw OMML output from `mathml_to_omml::convert` is structurally correct
// but needs several adjustments before PowerPoint will render it properly:
//
// 1. Cambria Math font injection in every run and structural properties block
// 2. Wrapping in `<m:oMathPara>` with center-group justification
// 3. Wrapping in `<a14:m>` for embedding inside `<a:p>` paragraphs

use regex::Regex;

/// The font-face attributes shared by every math run's `<a:rPr>`, regardless
/// of size — factored out so [`font_rpr_xml`]/[`ctrl_pr_xml`] only need to
/// decide whether an `sz=` attribute goes in front of it.
const FONT_FACE_ATTRS: &str =
    r#"typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0""#;

/// Build the `<a:rPr>` fragment PowerPoint expects in every math run, with an
/// optional explicit point size.
///
/// `sz_pt` is the surrounding text run's own font size — the same value a
/// sibling prose `<a:r><a:rPr sz="…">` in the same paragraph carries. Without
/// it, a math run's `<a:rPr>` carries no `sz` at all, and PowerPoint falls
/// back to its own built-in default (18pt) rather than inheriting the
/// paragraph's size the way a plain text run does — the "18pt OMML in 28pt
/// prose" bug (`wayfinder/deckgen`'s ticket: maths font size == surrounding
/// text size, always). `None` preserves the old sz-less behaviour, for
/// callers that don't know (or don't care about) the surrounding size.
fn font_rpr_xml(sz_pt: Option<f64>) -> String {
    match sz_pt {
        Some(pt) => {
            let sz = (pt * 100.0).round() as i64;
            format!(r#"<a:rPr sz="{sz}"><a:latin {FONT_FACE_ATTRS}/></a:rPr>"#)
        }
        None => format!(r#"<a:rPr><a:latin {FONT_FACE_ATTRS}/></a:rPr>"#),
    }
}

/// The `<m:ctrlPr>` element injected into structural property blocks, at the
/// same optional size as [`font_rpr_xml`] — control-property runs (fraction
/// bars, radical control points, etc.) need the same `sz` as the visible math
/// runs around them or PowerPoint's own size negotiation between the two can
/// disagree, most visibly on `\frac`'s bar thickness.
fn ctrl_pr_xml(sz_pt: Option<f64>) -> String {
    format!("<m:ctrlPr>{}</m:ctrlPr>", font_rpr_xml(sz_pt))
}

/// Apply all PowerPoint-specific post-processing to raw OMML output.
/// Returns the equation XML ready to embed inside a `<a:p>` paragraph.
/// Emits math runs with no explicit `sz` (PowerPoint's own default, 18pt) —
/// see [`postprocess_omml_sized`] for the font-parity-correct entry point.
pub fn postprocess_omml(raw_omml: &str) -> String {
    postprocess_omml_sized(raw_omml, None)
}

/// Same as [`postprocess_omml`], but every math run's `<a:rPr>` (and every
/// structural block's `<m:ctrlPr>`) carries an explicit `sz="…"` matching
/// `sz_pt` (the surrounding text run's own point size) when `Some`, so the
/// equation renders at font parity with the prose around it instead of
/// falling back to PowerPoint's built-in default.
pub fn postprocess_omml_sized(raw_omml: &str, sz_pt: Option<f64>) -> String {
    let s = inject_run_fonts(raw_omml, sz_pt);
    let s = inject_ctrl_pr(&s, sz_pt);
    let s = wrap_omath_para(&s);
    wrap_a14m(&s)
}

// ---------------------------------------------------------------------------
// Step 1a: Inject Cambria Math font into every <m:r> element
// ---------------------------------------------------------------------------

/// In every `<m:r>` run, ensure the font properties are set to Cambria Math
/// (and, when `sz_pt` is `Some`, an explicit matching size).
///
/// - If `<m:rPr>...</m:rPr>` exists, replace its contents with the font rPr.
/// - If no `<m:rPr>` exists, insert one right after `<m:r>`.
fn inject_run_fonts(omml: &str, sz_pt: Option<f64>) -> String {
    let font_rpr = font_rpr_xml(sz_pt);

    // First, replace existing <m:rPr>...</m:rPr> blocks inside <m:r> with our font.
    // The m:rPr content is simple (e.g. <m:sty m:val="..."/>) so a non-greedy match works.
    let re_rpr = Regex::new(r"<m:rPr>.*?</m:rPr>").unwrap();
    let with_replaced = re_rpr.replace_all(omml, font_rpr.as_str()).to_string();

    // Now handle <m:r> elements that had no <m:rPr> — insert one right after <m:r>.
    // Rust regex doesn't support lookahead, so we do this with simple string scanning.
    let tag = "<m:r>";
    let already = "<a:rPr";
    let mut result = String::with_capacity(with_replaced.len() + 256);
    let mut pos = 0;
    while pos < with_replaced.len() {
        if let Some(idx) = with_replaced[pos..].find(tag) {
            let abs = pos + idx;
            // Copy everything up to and including <m:r>
            result.push_str(&with_replaced[pos..abs + tag.len()]);
            let after = abs + tag.len();
            // Check if already followed by <a:rPr (open tag, with or without
            // attributes — the sized form now has a following space/attr
            // before its own `>`, so this can't just look for the literal
            // `<a:rPr>` substring the way the unsized-only version did).
            if !with_replaced[after..].starts_with(already) {
                result.push_str(&font_rpr);
            }
            pos = after;
        } else {
            result.push_str(&with_replaced[pos..]);
            break;
        }
    }

    result
}

// ---------------------------------------------------------------------------
// Step 1b: Inject <m:ctrlPr> into structural property blocks
// ---------------------------------------------------------------------------

/// The structural property element names that need a `<m:ctrlPr>` child.
const STRUCT_PR_TAGS: &[&str] = &[
    "m:fPr",
    "m:sSupPr",
    "m:sSubPr",
    "m:radPr",
    "m:dPr",
    "m:eqArrPr",
    "m:limLowPr",
    "m:limUppPr",
    "m:accPr",
    "m:barPr",
    "m:groupChrPr",
    "m:borderBoxPr",
    "m:funcPr",
    "m:phantPr",
    "m:sPrePr",
    "m:boxPr",
    "m:naryPr",
];

/// For each structural properties block (e.g. `<m:fPr>...</m:fPr>`), ensure it
/// contains a `<m:ctrlPr>` element (sized per `sz_pt`, see [`ctrl_pr_xml`]).
/// If one already exists, leave it alone.
fn inject_ctrl_pr(omml: &str, sz_pt: Option<f64>) -> String {
    let ctrl_pr = ctrl_pr_xml(sz_pt);
    let mut result = omml.to_string();

    for tag in STRUCT_PR_TAGS {
        let open = format!("<{}>", tag);
        let close = format!("</{}>", tag);

        // Find each occurrence of this properties block
        let mut search_from = 0;
        loop {
            let start = match result[search_from..].find(&open) {
                Some(pos) => search_from + pos,
                None => break,
            };
            let end = match result[start..].find(&close) {
                Some(pos) => start + pos,
                None => break,
            };

            let block = &result[start + open.len()..end];

            if !block.contains("<m:ctrlPr>") {
                // Insert <m:ctrlPr> just before the closing tag
                let insert_pos = end;
                result.insert_str(insert_pos, &ctrl_pr);
                // Advance past what we just inserted
                search_from = insert_pos + ctrl_pr.len() + close.len();
            } else {
                search_from = end + close.len();
            }
        }
    }

    result
}

// ---------------------------------------------------------------------------
// Step 2: Wrap in <m:oMathPara>
// ---------------------------------------------------------------------------

/// Wrap the `<m:oMath ...>` element in an `<m:oMathPara>` with center-group
/// justification, preserving the inner content.
fn wrap_omath_para(omml: &str) -> String {
    // Match <m:oMath ...>...</m:oMath> — the opening tag may have attributes/namespace.
    let re = Regex::new(r"(?s)<m:oMath(\s[^>]*)?>(.+)</m:oMath>").unwrap();

    if let Some(caps) = re.captures(omml) {
        let inner = &caps[2];
        format!(
            concat!(
                r#"<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">"#,
                r#"<m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>"#,
                r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">"#,
                "{}",
                "</m:oMath>",
                "</m:oMathPara>",
            ),
            inner
        )
    } else {
        // If we can't find the expected structure, return as-is
        omml.to_string()
    }
}

// ---------------------------------------------------------------------------
// Step 3: Wrap in <a14:m>
// ---------------------------------------------------------------------------

/// Wrap the entire equation in `<a14:m>...</a14:m>`.
fn wrap_a14m(omml: &str) -> String {
    format!("<a14:m>{}</a14:m>", omml)
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_run_font_injection_replaces_existing_rpr() {
        let input = r#"<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>x</m:t></m:r>"#;
        let result = inject_run_fonts(input, None);
        assert!(result.contains("Cambria Math"));
        assert!(!result.contains("<m:sty"));
    }

    #[test]
    fn test_run_font_injection_inserts_missing_rpr() {
        let input = "<m:r><m:t>x</m:t></m:r>";
        let result = inject_run_fonts(input, None);
        assert!(result.contains("Cambria Math"));
        assert!(result.contains("<a:rPr>"));
    }

    #[test]
    fn test_ctrl_pr_injection() {
        let input = "<m:fPr><m:type m:val=\"bar\"/></m:fPr>";
        let result = inject_ctrl_pr(input, None);
        assert!(result.contains("<m:ctrlPr>"));
        assert!(result.contains("Cambria Math"));
    }

    #[test]
    fn test_ctrl_pr_not_duplicated() {
        let input = "<m:fPr><m:ctrlPr><a:rPr/></m:ctrlPr></m:fPr>";
        let result = inject_ctrl_pr(input, None);
        // Should still have exactly one ctrlPr
        assert_eq!(result.matches("<m:ctrlPr>").count(), 1);
    }

    // ── size-parity injection (sz_pt threaded through) ──────────────────────

    #[test]
    fn run_font_injection_carries_explicit_size_when_given() {
        let input = "<m:r><m:t>x</m:t></m:r>";
        let result = inject_run_fonts(input, Some(28.0));
        assert!(
            result.contains(r#"<a:rPr sz="2800">"#),
            "expected sz=\"2800\" (28pt in hundredths-of-a-point, matching \
             `<a:rPr sz=…>` on an ordinary text run): {result}"
        );
        assert!(result.contains("Cambria Math"));
    }

    #[test]
    fn run_font_injection_replaces_existing_rpr_with_sized_one() {
        let input = r#"<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>x</m:t></m:r>"#;
        let result = inject_run_fonts(input, Some(44.0));
        assert!(result.contains(r#"sz="4400""#));
        assert!(!result.contains("<m:sty"));
    }

    #[test]
    fn run_font_injection_does_not_double_inject_a_sized_rpr() {
        // Regression for the `already.starts_with("<a:rPr>")` exact-match
        // check, which would miss a sized `<a:rPr sz="…">` (no immediate `>`)
        // and duplicate the font rPr on every subsequent <m:r>.
        let input = "<m:r><m:t>x</m:t></m:r><m:r><m:t>y</m:t></m:r>";
        let result = inject_run_fonts(input, Some(28.0));
        assert_eq!(
            result.matches("Cambria Math").count(),
            2,
            "each <m:r> gets exactly one font rPr, not a duplicate: {result}"
        );
    }

    #[test]
    fn ctrl_pr_injection_carries_explicit_size_when_given() {
        let input = "<m:fPr><m:type m:val=\"bar\"/></m:fPr>";
        let result = inject_ctrl_pr(input, Some(28.0));
        assert!(result.contains(r#"sz="2800""#));
    }

    #[test]
    fn test_omath_para_wrapping() {
        let input = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>"#;
        let result = wrap_omath_para(input);
        assert!(result.contains("<m:oMathPara"));
        assert!(result.contains("<m:oMathParaPr>"));
        assert!(result.contains("centerGroup"));
    }

    #[test]
    fn test_a14m_wrapping() {
        let input = "<m:oMathPara>...</m:oMathPara>";
        let result = wrap_a14m(input);
        assert!(result.starts_with("<a14:m>"));
        assert!(result.ends_with("</a14:m>"));
    }

    #[test]
    fn test_postprocess_full_pipeline() {
        let input = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>"#;
        let result = postprocess_omml(input);
        assert!(result.contains("<a14:m>"));
        assert!(result.contains("Cambria Math"));
        assert!(result.contains("<m:oMathPara"));
        assert!(result.contains("</a14:m>"));
        assert!(
            !result.contains("sz="),
            "the unsized entry point must not invent a size: {result}"
        );
    }

    /// Font-parity regression: `postprocess_omml_sized(_, Some(28.0))` must
    /// put `sz="2800"` on every math run's `<a:rPr>`, matching a sibling
    /// prose run's own `sz="2800"` at 28pt — the "18pt OMML in 28pt prose"
    /// bug this threading exists to fix.
    #[test]
    fn test_postprocess_sized_pipeline_matches_surrounding_text_size() {
        let input = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r><m:r><m:t>y</m:t></m:r></m:oMath>"#;
        let result = postprocess_omml_sized(input, Some(28.0));
        assert!(result.contains("<a14:m>"));
        assert_eq!(
            result.matches(r#"sz="2800""#).count(),
            2,
            "both math runs must carry sz=\"2800\": {result}"
        );
    }
}
