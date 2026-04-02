pub mod mathml_to_omml;
pub mod postprocess;

pub use mathml_to_omml::MathError;

/// Convert a LaTeX equation string to PowerPoint-ready OMML XML.
///
/// Returns the complete `<a14:m>` element ready to embed in a `<a:p>` paragraph.
pub fn latex_to_omml(latex: &str) -> Result<String, MathError> {
    let mathml = latex2mathml::latex_to_mathml(latex, latex2mathml::DisplayStyle::Inline)
        .map_err(|e| MathError::Latex(e.to_string()))?;
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
}
