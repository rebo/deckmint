// postprocess.rs — PowerPoint-specific post-processing for OMML equations
//
// The raw OMML output from `mathml_to_omml::convert` is structurally correct
// but needs several adjustments before PowerPoint will render it properly:
//
// 1. Cambria Math font injection in every run and structural properties block
// 2. Wrapping in `<m:oMathPara>` with center-group justification
// 3. Wrapping in `<a14:m>` for embedding inside `<a:p>` paragraphs

use regex::Regex;

/// The font `<a:rPr>` fragment that PowerPoint expects in every math run.
const FONT_RPR: &str = r#"<a:rPr><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>"#;

/// The `<m:ctrlPr>` element injected into structural property blocks.
const CTRL_PR: &str = concat!(
    r#"<m:ctrlPr><a:rPr><a:latin typeface="Cambria Math" "#,
    r#"panose="02040503050406030204" pitchFamily="18" charset="0"/>"#,
    r#"</a:rPr></m:ctrlPr>"#,
);

/// Apply all PowerPoint-specific post-processing to raw OMML output.
/// Returns the equation XML ready to embed inside a `<a:p>` paragraph.
pub fn postprocess_omml(raw_omml: &str) -> String {
    let s = inject_run_fonts(raw_omml);
    let s = inject_ctrl_pr(&s);
    let s = wrap_omath_para(&s);
    wrap_a14m(&s)
}

// ---------------------------------------------------------------------------
// Step 1a: Inject Cambria Math font into every <m:r> element
// ---------------------------------------------------------------------------

/// In every `<m:r>` run, ensure the font properties are set to Cambria Math.
///
/// - If `<m:rPr>...</m:rPr>` exists, replace its contents with the font rPr.
/// - If no `<m:rPr>` exists, insert one right after `<m:r>`.
fn inject_run_fonts(omml: &str) -> String {
    // First, replace existing <m:rPr>...</m:rPr> blocks inside <m:r> with our font.
    // The m:rPr content is simple (e.g. <m:sty m:val="..."/>) so a non-greedy match works.
    let re_rpr = Regex::new(r"<m:rPr>.*?</m:rPr>").unwrap();
    let with_replaced = re_rpr.replace_all(omml, FONT_RPR).to_string();

    // Now handle <m:r> elements that had no <m:rPr> — insert one right after <m:r>.
    // Rust regex doesn't support lookahead, so we do this with simple string scanning.
    let tag = "<m:r>";
    let already = "<a:rPr>";
    let mut result = String::with_capacity(with_replaced.len() + 256);
    let mut pos = 0;
    while pos < with_replaced.len() {
        if let Some(idx) = with_replaced[pos..].find(tag) {
            let abs = pos + idx;
            // Copy everything up to and including <m:r>
            result.push_str(&with_replaced[pos..abs + tag.len()]);
            let after = abs + tag.len();
            // Check if already followed by <a:rPr>
            if !with_replaced[after..].starts_with(already) {
                result.push_str(FONT_RPR);
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
/// contains a `<m:ctrlPr>` element. If one already exists, leave it alone.
fn inject_ctrl_pr(omml: &str) -> String {
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
                result.insert_str(insert_pos, CTRL_PR);
                // Advance past what we just inserted
                search_from = insert_pos + CTRL_PR.len() + close.len();
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
        let result = inject_run_fonts(input);
        assert!(result.contains("Cambria Math"));
        assert!(!result.contains("<m:sty"));
    }

    #[test]
    fn test_run_font_injection_inserts_missing_rpr() {
        let input = "<m:r><m:t>x</m:t></m:r>";
        let result = inject_run_fonts(input);
        assert!(result.contains("Cambria Math"));
        assert!(result.contains("<a:rPr>"));
    }

    #[test]
    fn test_ctrl_pr_injection() {
        let input = "<m:fPr><m:type m:val=\"bar\"/></m:fPr>";
        let result = inject_ctrl_pr(input);
        assert!(result.contains("<m:ctrlPr>"));
        assert!(result.contains("Cambria Math"));
    }

    #[test]
    fn test_ctrl_pr_not_duplicated() {
        let input = "<m:fPr><m:ctrlPr><a:rPr/></m:ctrlPr></m:fPr>";
        let result = inject_ctrl_pr(input);
        // Should still have exactly one ctrlPr
        assert_eq!(result.matches("<m:ctrlPr>").count(), 1);
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
    }
}
