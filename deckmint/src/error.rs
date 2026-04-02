use std::fmt;

/// All errors produced by the deckmint library
#[derive(Debug)]
pub enum PptxError {
    Zip(zip::result::ZipError),
    Io(std::io::Error),
    InvalidColor(String),
    InvalidCoord(String),
    InvalidArgument(String),
    Base64(String),
    /// Structural validation errors found in the generated PPTX.
    ValidationFailed(Vec<crate::validate::ValidationIssue>),
}

impl fmt::Display for PptxError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            PptxError::Zip(e) => write!(f, "ZIP error: {e}"),
            PptxError::Io(e) => write!(f, "I/O error: {e}"),
            PptxError::InvalidColor(s) => write!(f, "Invalid color: {s}"),
            PptxError::InvalidCoord(s) => write!(f, "Invalid coordinate: {s}"),
            PptxError::InvalidArgument(s) => write!(f, "Invalid argument: {s}"),
            PptxError::Base64(s) => write!(f, "Base64 error: {s}"),
            PptxError::ValidationFailed(issues) => {
                write!(f, "Validation failed with {} issue(s):", issues.len())?;
                for issue in issues {
                    write!(f, "\n  [{:?}] {}", issue.severity, issue.message)?;
                    if let Some(ref part) = issue.part {
                        write!(f, " (in {part})")?;
                    }
                }
                Ok(())
            }
        }
    }
}

impl std::error::Error for PptxError {}

impl From<zip::result::ZipError> for PptxError {
    fn from(e: zip::result::ZipError) -> Self {
        PptxError::Zip(e)
    }
}

impl From<std::io::Error> for PptxError {
    fn from(e: std::io::Error) -> Self {
        PptxError::Io(e)
    }
}
