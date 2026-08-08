pub mod mathml_to_omml;
pub mod postprocess;
mod xml_repair;

pub use mathml_to_omml::MathError;

/// Convert a LaTeX equation string to PowerPoint-ready OMML XML.
///
/// Returns the complete `<a14:m>` element ready to embed in a `<a:p>` paragraph.
///
/// Before parsing, the intermediate MathML from `latex2mathml` (0.2.3) is
/// repaired for two known serialisation bugs in that crate — see
/// `xml_repair`'s module docs — so this is safe for any input that
/// `latex2mathml` itself tokenises successfully, including LaTeX containing a
/// literal `<`/`&` (e.g. `x < 3`) or a `\bigl(...\bigr)` sized-paren.
pub fn latex_to_omml(latex: &str) -> Result<String, MathError> {
    let mathml = latex2mathml::latex_to_mathml(latex, latex2mathml::DisplayStyle::Inline)
        .map_err(|e| MathError::Latex(e.to_string()))?;
    let mathml = xml_repair::fix_sizedparen_typo(&xml_repair::escape_bad_xml_chars(&mathml));
    let omml = mathml_to_omml::convert(&mathml)?;
    Ok(postprocess::postprocess_omml(&omml))
}

/// Convert a MathML string to PowerPoint-ready OMML XML.
pub fn mathml_to_pptx_omml(mathml: &str) -> Result<String, MathError> {
    let omml = mathml_to_omml::convert(mathml)?;
    Ok(postprocess::postprocess_omml(&omml))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_latex_to_omml_basic() {
        let result = latex_to_omml(r"x^2 + 3x - 7").unwrap();
        assert!(result.contains("<a14:m>"));
        assert!(result.contains("Cambria Math"));
        assert!(result.contains("<m:oMathPara"));
    }

    // ── regression tests for the two `latex2mathml` 0.2.3 serialisation bugs
    // (maths-sweep.md categories 8–10, from the revision_atlas deckgen corpus
    // sweep) ─────────────────────────────────────────────────────────────

    /// Category 8: a literal strict-inequality `<` tokenises fine but
    /// `Node::Operator`'s `Display` impl writes it unescaped into MathML text
    /// content (`<mo><</mo>`), which is invalid XML until `xml_repair`
    /// escapes it.
    #[test]
    fn literal_strict_inequality_converts_cleanly() {
        let result = latex_to_omml(r"x < 3").unwrap_or_else(|e| panic!("x < 3 failed to convert: {e}"));
        assert!(result.contains("<a14:m>"));
    }

    /// Category 9: same unescaped-serialisation bug, via `Node::Letter` for a
    /// literal `\&`.
    #[test]
    fn literal_ampersand_converts_cleanly() {
        let result = latex_to_omml(r"A \& B").unwrap_or_else(|e| panic!(r"A \& B failed to convert: {e}"));
        assert!(result.contains("<a14:m>"));
    }

    /// Category 10: `Node::SizedParen`'s `Display` impl closes `<mo …>` with a
    /// literal `</mro>` typo instead of `</mo>`. `\bigl`/`\bigr` are already
    /// recognised tokens in latex2mathml 0.2.3 (unlike bare `\big`, which
    /// needs a text-level rewrite the caller is responsible for — deckmint-math
    /// only owns the MathML-level repair), so this exercises the bug directly.
    #[test]
    fn sized_paren_converts_cleanly() {
        let result = latex_to_omml(r"\bigl(\frac{a}{b}\bigr)")
            .unwrap_or_else(|e| panic!(r"\bigl(\frac{{a}}{{b}}\bigr) failed to convert: {e}"));
        assert!(result.contains("<a14:m>"));
    }
}
