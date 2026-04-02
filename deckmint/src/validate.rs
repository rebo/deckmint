//! Post-generation validation for PPTX archives.
//!
//! Validates structural integrity of generated PPTX files. In a generation-only
//! library, validation failures indicate bugs in the generators — not user data
//! issues. Use [`validate`] to check output, or [`Presentation::write_validated`]
//! for integrated checking.

use std::collections::HashSet;
use std::io::Read;

/// Severity of a validation issue.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Severity {
    /// Critical structural problem — file may not open.
    Error,
    /// Non-critical issue — file opens but may behave unexpectedly.
    Warning,
}

/// A single validation issue found in the PPTX archive.
#[derive(Debug, Clone)]
pub struct ValidationIssue {
    /// Severity level.
    pub severity: Severity,
    /// Human-readable description of the issue.
    pub message: String,
    /// ZIP entry path where the issue was found (if applicable).
    pub part: Option<String>,
}

/// Validate a PPTX archive from its raw bytes.
///
/// Returns a list of issues found. An empty list means the archive passes
/// all structural checks.
pub fn validate(pptx_bytes: &[u8]) -> Vec<ValidationIssue> {
    let mut issues = Vec::new();

    let cursor = std::io::Cursor::new(pptx_bytes);
    let mut archive = match zip::ZipArchive::new(cursor) {
        Ok(a) => a,
        Err(e) => {
            issues.push(ValidationIssue {
                severity: Severity::Error,
                message: format!("Cannot open PPTX as ZIP: {e}"),
                part: None,
            });
            return issues;
        }
    };

    // Collect all entry names
    let entry_names: HashSet<String> = (0..archive.len())
        .filter_map(|i| archive.by_index(i).ok().map(|f| f.name().to_string()))
        .collect();

    // Check 1: presentation.xml exists
    check_presentation_part(&entry_names, &mut issues);

    // Check 2: root rels valid
    check_root_rels(&entry_names, &mut archive, &mut issues);

    // Check 3: content types complete
    check_content_types(&entry_names, &mut archive, &mut issues);

    // Check 4: XML well-formedness (basic check)
    check_xml_wellformedness(&entry_names, &mut archive, &mut issues);

    // Check 5: relationship targets resolve
    check_relationship_targets(&entry_names, &mut archive, &mut issues);

    issues
}

fn check_presentation_part(entries: &HashSet<String>, issues: &mut Vec<ValidationIssue>) {
    if !entries.contains("ppt/presentation.xml") {
        issues.push(ValidationIssue {
            severity: Severity::Error,
            message: "Missing ppt/presentation.xml".to_string(),
            part: Some("ppt/presentation.xml".to_string()),
        });
    }
}

fn check_root_rels(entries: &HashSet<String>, archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>, issues: &mut Vec<ValidationIssue>) {
    if !entries.contains("_rels/.rels") {
        issues.push(ValidationIssue {
            severity: Severity::Error,
            message: "Missing _rels/.rels".to_string(),
            part: Some("_rels/.rels".to_string()),
        });
        return;
    }

    if let Ok(mut file) = archive.by_name("_rels/.rels") {
        let mut content = String::new();
        if file.read_to_string(&mut content).is_ok() {
            if !content.contains("officeDocument") {
                issues.push(ValidationIssue {
                    severity: Severity::Error,
                    message: "_rels/.rels missing officeDocument relationship".to_string(),
                    part: Some("_rels/.rels".to_string()),
                });
            }
        }
    }
}

fn check_content_types(entries: &HashSet<String>, archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>, issues: &mut Vec<ValidationIssue>) {
    if !entries.contains("[Content_Types].xml") {
        issues.push(ValidationIssue {
            severity: Severity::Error,
            message: "Missing [Content_Types].xml".to_string(),
            part: Some("[Content_Types].xml".to_string()),
        });
        return;
    }

    if let Ok(mut file) = archive.by_name("[Content_Types].xml") {
        let mut content = String::new();
        if file.read_to_string(&mut content).is_ok() {
            // Check that all XML parts are covered
            for entry in entries {
                if entry.ends_with(".xml") && entry != "[Content_Types].xml" && !entry.contains("_rels/") {
                    let part_name = format!("/{entry}");
                    if !content.contains(&part_name) && !content.contains("Extension=\"xml\"") {
                        issues.push(ValidationIssue {
                            severity: Severity::Warning,
                            message: format!("Part {entry} not covered by [Content_Types].xml"),
                            part: Some(entry.clone()),
                        });
                    }
                }
            }
        }
    }
}

fn check_xml_wellformedness(entries: &HashSet<String>, archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>, issues: &mut Vec<ValidationIssue>) {
    for entry_name in entries {
        if !entry_name.ends_with(".xml") && !entry_name.ends_with(".rels") {
            continue;
        }
        if let Ok(mut file) = archive.by_name(entry_name) {
            let mut content = String::new();
            if file.read_to_string(&mut content).is_ok() {
                // Basic well-formedness: check balanced angle brackets
                let opens: usize = content.matches('<').count();
                let closes: usize = content.matches('>').count();
                if opens != closes {
                    issues.push(ValidationIssue {
                        severity: Severity::Error,
                        message: format!("Malformed XML: unbalanced angle brackets ({opens} '<' vs {closes} '>')"),
                        part: Some(entry_name.clone()),
                    });
                }
            }
        }
    }
}

fn check_relationship_targets(entries: &HashSet<String>, archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>, issues: &mut Vec<ValidationIssue>) {
    let rels_files: Vec<String> = entries.iter()
        .filter(|e| e.ends_with(".rels"))
        .cloned()
        .collect();

    for rels_file in &rels_files {
        if let Ok(mut file) = archive.by_name(rels_file) {
            let mut content = String::new();
            if file.read_to_string(&mut content).is_err() {
                continue;
            }

            // Extract Target attributes from <Relationship> elements
            // Skip External rels (TargetMode="External")
            for line in content.split("<Relationship ") {
                if line.contains("TargetMode=\"External\"") {
                    continue;
                }
                if let Some(target) = extract_attr(line, "Target") {
                    if target.is_empty() || target.starts_with("http://") || target.starts_with("https://") {
                        continue;
                    }
                    // Resolve relative path
                    let base_dir = rels_file.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
                    // rels files are in _rels/ subdirs, targets are relative to the parent
                    let parent_dir = base_dir.replace("_rels", "").replace("//", "/");
                    let resolved = resolve_rel_path(&parent_dir, &target);
                    if !entries.contains(&resolved) {
                        issues.push(ValidationIssue {
                            severity: Severity::Error,
                            message: format!("Broken relationship: target '{target}' (resolved to '{resolved}') does not exist in archive"),
                            part: Some(rels_file.clone()),
                        });
                    }
                }
            }
        }
    }
}

/// Extract an XML attribute value from a string fragment.
fn extract_attr(s: &str, attr: &str) -> Option<String> {
    let pattern = format!("{attr}=\"");
    if let Some(start) = s.find(&pattern) {
        let rest = &s[start + pattern.len()..];
        if let Some(end) = rest.find('"') {
            return Some(rest[..end].to_string());
        }
    }
    None
}

/// Resolve a relative path against a base directory.
fn resolve_rel_path(base: &str, target: &str) -> String {
    if target.starts_with('/') {
        return target.trim_start_matches('/').to_string();
    }
    let mut parts: Vec<&str> = base.split('/').filter(|s| !s.is_empty()).collect();
    for segment in target.split('/') {
        match segment {
            ".." => { parts.pop(); }
            "." | "" => {}
            s => parts.push(s),
        }
    }
    parts.join("/")
}
