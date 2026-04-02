use crate::objects::{SlideRel, SlideRelMedia};
use crate::xml::CRLF;

// ─────────────────────────────────────────────────────────────
// ppt/_rels/presentation.xml.rels
// ─────────────────────────────────────────────────────────────

pub fn make_xml_presentation_rels(slide_count: usize) -> String {
    let mut s = String::new();
    s.push_str(&format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{CRLF}\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
    ));
    s.push_str("<Relationship Id=\"rId1\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" \
Target=\"slideMasters/slideMaster1.xml\"/>");
    for idx in 1..=slide_count {
        let rid = idx + 1;
        s.push_str(&format!(
            "<Relationship Id=\"rId{rid}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" \
Target=\"slides/slide{idx}.xml\"/>"
        ));
    }
    let base = slide_count + 2; // 1 (master) + slide_count + 1
    s.push_str(&format!(
        "<Relationship Id=\"rId{n}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster\" \
Target=\"notesMasters/notesMaster1.xml\"/>",
        n = base
    ));
    s.push_str(&format!(
        "<Relationship Id=\"rId{n}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps\" \
Target=\"presProps.xml\"/>",
        n = base + 1
    ));
    s.push_str(&format!(
        "<Relationship Id=\"rId{n}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps\" \
Target=\"viewProps.xml\"/>",
        n = base + 2
    ));
    s.push_str(&format!(
        "<Relationship Id=\"rId{n}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" \
Target=\"theme/theme1.xml\"/>",
        n = base + 3
    ));
    s.push_str(&format!(
        "<Relationship Id=\"rId{n}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles\" \
Target=\"tableStyles.xml\"/>",
        n = base + 4
    ));
    s.push_str("</Relationships>");
    s
}

// ─────────────────────────────────────────────────────────────
// ppt/slides/_rels/slideN.xml.rels
// ─────────────────────────────────────────────────────────────

pub fn make_xml_slide_rel(
    slide_num: usize,
    layout_idx: usize,
    rels: &[SlideRel],
    rels_media: &[SlideRelMedia],
    media_targets: &[String],
    chart_rels: &[(u32, String)],  // (chart_rid, chart_target) pairs
) -> String {
    let mut s = String::new();
    s.push_str(&format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{CRLF}\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
    ));

    // rId1 = slideLayout, rId2 = notesSlide (always first, as PowerPoint expects)
    s.push_str(&format!(
        "<Relationship Id=\"rId1\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" \
Target=\"../slideLayouts/slideLayout{layout_idx}.xml\"/>"
    ));
    s.push_str(&format!(
        "<Relationship Id=\"rId2\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide\" \
Target=\"../notesSlides/notesSlide{slide_num}.xml\"/>"
    ));

    // User content rels (hyperlinks, slide-jumps) — start at rId3+
    for rel in rels {
        if rel.rel_type.to_lowercase().contains("hyperlink") {
            if rel.data.as_deref() == Some("slide") {
                s.push_str(&format!(
                    "<Relationship Id=\"rId{}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" \
Target=\"slide{}.xml\"/>",
                    rel.r_id, rel.target
                ));
            } else {
                s.push_str(&format!(
                    "<Relationship Id=\"rId{}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" \
Target=\"{}\" TargetMode=\"External\"/>",
                    rel.r_id, rel.target
                ));
            }
        }
    }

    // Media rels (images, video, audio) — start at rId3+, use global sequential targets
    for (i, rel) in rels_media.iter().enumerate() {
        let target = media_targets.get(i).map(String::as_str).unwrap_or(&rel.target);
        let type_uri = match rel.rel_type.as_str() {
            "video" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
            "audio" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
            _ => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        };
        s.push_str(&format!(
            "<Relationship Id=\"rId{}\" \
Type=\"{type_uri}\" \
Target=\"{target}\"/>",
            rel.r_id
        ));
    }

    // Chart rels
    for (rid, target) in chart_rels {
        s.push_str(&format!(
            "<Relationship Id=\"rId{}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" \
Target=\"{target}\"/>",
            rid
        ));
    }

    s.push_str("</Relationships>");
    s
}

// ─────────────────────────────────────────────────────────────
// ppt/slideLayouts/_rels/slideLayout1.xml.rels  (static — one layout)
// ─────────────────────────────────────────────────────────────

pub fn make_xml_slide_layout_rel() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" \
Target=\"../slideMasters/slideMaster1.xml\"/>\
</Relationships>"
}

// ─────────────────────────────────────────────────────────────
// ppt/slideMasters/_rels/slideMaster1.xml.rels
// ─────────────────────────────────────────────────────────────

pub fn make_xml_master_rel(
    layout_count: usize,
    rels: &[SlideRel],
    rels_media: &[SlideRelMedia],
    media_targets: &[String],
    rid_offset: u32,
) -> String {
    let mut s = String::new();
    s.push_str(&format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{CRLF}\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
    ));
    for idx in 1..=layout_count {
        s.push_str(&format!(
            "<Relationship Id=\"rId{idx}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" \
Target=\"../slideLayouts/slideLayout{idx}.xml\"/>"
        ));
    }
    let theme_rid = layout_count + 1;
    s.push_str(&format!(
        "<Relationship Id=\"rId{theme_rid}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" \
Target=\"../theme/theme1.xml\"/>"
    ));
    // Hyperlink rels
    for rel in rels {
        let actual_rid = rel.r_id + rid_offset;
        if rel.rel_type.to_lowercase().contains("hyperlink") {
            if rel.data.as_deref() == Some("slide") {
                s.push_str(&format!(
                    "<Relationship Id=\"rId{actual_rid}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" \
Target=\"slide{}.xml\"/>", rel.target));
            } else {
                s.push_str(&format!(
                    "<Relationship Id=\"rId{actual_rid}\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" \
Target=\"{}\" TargetMode=\"External\"/>", rel.target));
            }
        }
    }
    // Media rels (images, etc.)
    for (i, rel) in rels_media.iter().enumerate() {
        let actual_rid = rel.r_id + rid_offset;
        let target = media_targets.get(i).map(String::as_str).unwrap_or(&rel.target);
        let type_uri = match rel.rel_type.as_str() {
            "video" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
            "audio" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
            _ => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        };
        s.push_str(&format!(
            "<Relationship Id=\"rId{actual_rid}\" Type=\"{type_uri}\" Target=\"{target}\"/>"));
    }
    s.push_str("</Relationships>");
    s
}

// ─────────────────────────────────────────────────────────────
// ppt/notesSlides/_rels/notesSlideN.xml.rels
// ─────────────────────────────────────────────────────────────

pub fn make_xml_notes_slide_rel(slide_num: usize) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{CRLF}\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster\" \
Target=\"../notesMasters/notesMaster1.xml\"/>\
<Relationship Id=\"rId2\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" \
Target=\"../slides/slide{slide_num}.xml\"/>\
</Relationships>"
    )
}

// ─────────────────────────────────────────────────────────────
// ppt/notesMasters/_rels/notesMaster1.xml.rels  (static)
// ─────────────────────────────────────────────────────────────

pub fn make_xml_notes_master_rel() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" \
Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" \
Target=\"../theme/theme2.xml\"/>\
</Relationships>"
}
