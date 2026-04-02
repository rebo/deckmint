use crate::presentation::{Presentation, SlideLayout};
use crate::slide::Slide;
use crate::utils::encode_xml_entities;
use crate::xml::CRLF;

// ─────────────────────────────────────────────────────────────
// [Content_Types].xml
// ─────────────────────────────────────────────────────────────

pub fn make_xml_content_types(pres: &Presentation) -> String {
    let mut s = String::new();
    s.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    s.push_str(CRLF);
    s.push_str("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
    // Always-required defaults
    s.push_str("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
    s.push_str("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
    // Only declare image MIME types for extensions actually present in this package
    let mime_map: &[(&str, &str)] = &[
        ("jpeg", "image/jpeg"),
        ("jpg",  "image/jpeg"),
        ("png",  "image/png"),
        ("gif",  "image/gif"),
        ("svg",  "image/svg+xml"),
        ("webp", "image/webp"),
        ("bmp",  "image/bmp"),
        ("tiff", "image/tiff"),
        ("tif",  "image/tiff"),
        ("emf",  "image/x-emf"),
        ("wmf",  "image/x-wmf"),
        ("mp4",  "video/mp4"),
        ("mov",  "video/quicktime"),
        ("avi",  "video/x-msvideo"),
        ("wmv",  "video/x-ms-wmv"),
        ("mp3",  "audio/mpeg"),
        ("wav",  "audio/wav"),
        ("m4a",  "audio/mp4"),
    ];
    let mut used_extns: std::collections::HashSet<&str> = pres.slides.iter()
        .flat_map(|sl| sl.rels_media.iter().map(|m| m.extn.as_str()))
        .collect();
    if let Some(ref m) = pres.master {
        for media in &m.rels_media {
            used_extns.insert(&media.extn);
        }
    }
    for (ext, mime) in mime_map {
        if used_extns.contains(ext) {
            s.push_str(&format!("<Default Extension=\"{ext}\" ContentType=\"{mime}\"/>"));
        }
    }
    // Presentation
    s.push_str("<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>");
    s.push_str("<Override PartName=\"/ppt/notesMasters/notesMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml\"/>");
    // Slide master
    s.push_str("<Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>");
    // Slides
    for (idx, _slide) in pres.slides.iter().enumerate() {
        s.push_str(&format!(
            "<Override PartName=\"/ppt/slides/slide{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>",
            idx + 1
        ));
    }
    // Notes slides
    for (idx, _) in pres.slides.iter().enumerate() {
        s.push_str(&format!(
            "<Override PartName=\"/ppt/notesSlides/notesSlide{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml\"/>",
            idx + 1
        ));
    }
    // Slide layouts
    for (idx, _) in pres.slide_layouts.iter().enumerate() {
        s.push_str(&format!(
            "<Override PartName=\"/ppt/slideLayouts/slideLayout{}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>",
            idx + 1
        ));
    }
    // Charts
    let mut chart_idx = 0usize;
    for slide in &pres.slides {
        for _ in &slide.charts {
            chart_idx += 1;
            s.push_str(&format!(
                "<Override PartName=\"/ppt/charts/chart{chart_idx}.xml\" \
ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>"
            ));
        }
    }
    // Core PPT files
    s.push_str("<Override PartName=\"/ppt/presProps.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presProps+xml\"/>");
    s.push_str("<Override PartName=\"/ppt/viewProps.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml\"/>");
    s.push_str("<Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
    s.push_str("<Override PartName=\"/ppt/theme/theme2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
    s.push_str("<Override PartName=\"/ppt/tableStyles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml\"/>");
    s.push_str("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
    s.push_str("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
    s.push_str("</Types>");
    s
}

// ─────────────────────────────────────────────────────────────
// _rels/.rels
// ─────────────────────────────────────────────────────────────

pub fn make_xml_root_rels() -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{CRLF}\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>\
<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\
<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\
</Relationships>"
    )
}

// ─────────────────────────────────────────────────────────────
// docProps/app.xml
// ─────────────────────────────────────────────────────────────

pub fn make_xml_app(slides: &[Slide], company: &str) -> String {
    let slide_count = slides.len();
    let slide_titles: String = slides
        .iter()
        .enumerate()
        .map(|(i, _)| format!("<vt:lpstr>Slide {}</vt:lpstr>", i + 1))
        .collect();
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{CRLF}\
<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" \
xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\
<TotalTime>0</TotalTime><Words>0</Words>\
<Application>Microsoft Office PowerPoint</Application>\
<PresentationFormat>On-screen Show (16:9)</PresentationFormat>\
<Paragraphs>0</Paragraphs>\
<Slides>{slide_count}</Slides>\
<Notes>{slide_count}</Notes>\
<HiddenSlides>0</HiddenSlides>\
<MMClips>0</MMClips>\
<ScaleCrop>false</ScaleCrop>\
<HeadingPairs>\
<vt:vector size=\"6\" baseType=\"variant\">\
<vt:variant><vt:lpstr>Fonts Used</vt:lpstr></vt:variant>\
<vt:variant><vt:i4>2</vt:i4></vt:variant>\
<vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>\
<vt:variant><vt:i4>1</vt:i4></vt:variant>\
<vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>\
<vt:variant><vt:i4>{slide_count}</vt:i4></vt:variant>\
</vt:vector>\
</HeadingPairs>\
<TitlesOfParts>\
<vt:vector size=\"{total_parts}\" baseType=\"lpstr\">\
<vt:lpstr>Arial</vt:lpstr>\
<vt:lpstr>Calibri</vt:lpstr>\
<vt:lpstr>Office Theme</vt:lpstr>\
{slide_titles}\
</vt:vector>\
</TitlesOfParts>\
<Company>{company_enc}</Company>\
<LinksUpToDate>false</LinksUpToDate>\
<SharedDoc>false</SharedDoc>\
<HyperlinksChanged>false</HyperlinksChanged>\
<AppVersion>16.0000</AppVersion>\
</Properties>",
        slide_count = slide_count,
        total_parts = slide_count + 3,
        slide_titles = slide_titles,
        company_enc = encode_xml_entities(company),
    )
}

// ─────────────────────────────────────────────────────────────
// docProps/core.xml
// ─────────────────────────────────────────────────────────────

pub fn make_xml_core(title: &str, subject: &str, author: &str, revision: &str) -> String {
    let ts = iso_timestamp();
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{CRLF}\
<cp:coreProperties \
xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" \
xmlns:dc=\"http://purl.org/dc/elements/1.1/\" \
xmlns:dcterms=\"http://purl.org/dc/terms/\" \
xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" \
xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\
<dc:title>{title}</dc:title>\
<dc:subject>{subject}</dc:subject>\
<dc:creator>{author}</dc:creator>\
<cp:lastModifiedBy>{author}</cp:lastModifiedBy>\
<cp:revision>{revision}</cp:revision>\
<dcterms:created xsi:type=\"dcterms:W3CDTF\">{ts}</dcterms:created>\
<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{ts}</dcterms:modified>\
</cp:coreProperties>",
        title = encode_xml_entities(title),
        subject = encode_xml_entities(subject),
        author = encode_xml_entities(author),
        revision = revision,
        ts = ts,
    )
}

fn iso_timestamp() -> String {
    #[cfg(target_arch = "wasm32")]
    {
        // Use JS Date in WASM
        // This requires js-sys, which is only available in deckmint-wasm crate.
        // For the core crate on WASM without js-sys, fall back to a static string.
        "2024-01-01T00:00:00Z".to_string()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        use std::time::{SystemTime, UNIX_EPOCH};
        let secs = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default()
            .as_secs();
        // Format as ISO 8601 UTC without external dependency
        let (y, mo, d, h, mi, s) = seconds_to_datetime(secs);
        format!("{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z", y, mo, d, h, mi, s)
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn seconds_to_datetime(secs: u64) -> (u64, u64, u64, u64, u64, u64) {
    let s = secs % 60;
    let mins = secs / 60;
    let mi = mins % 60;
    let hours = mins / 60;
    let h = hours % 24;
    let days = hours / 24;
    // Approximate: good enough for a timestamp in core.xml
    let year = 1970 + days / 365;
    let day_of_year = days % 365;
    let month = (day_of_year / 30).min(11) + 1;
    let day = (day_of_year % 30) + 1;
    (year, month, day, h, mi, s)
}

// ─────────────────────────────────────────────────────────────
// ppt/presentation.xml
// ─────────────────────────────────────────────────────────────

pub fn make_xml_presentation(pres: &Presentation) -> String {
    let rtl_attr = if pres.rtl_mode { " rtl=\"1\"" } else { "" };
    let mut s = String::new();
    s.push_str(&format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{CRLF}\
<p:presentation \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \
xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"{rtl_attr} \
saveSubsetFonts=\"1\" autoCompressPictures=\"0\">"
    ));
    // Slide master
    s.push_str("<p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>");
    // Notes master (rId = slide_count + 2) — must precede sldIdLst per schema
    s.push_str(&format!(
        "<p:notesMasterIdLst><p:notesMasterId r:id=\"rId{}\"/></p:notesMasterIdLst>",
        pres.slides.len() + 2
    ));
    // Slides — rIds are index-based (rId2 = slide 1, rId3 = slide 2, etc.)
    // to stay in sync with presentation.xml.rels regardless of slide removal.
    s.push_str("<p:sldIdLst>");
    for (idx, slide) in pres.slides.iter().enumerate() {
        let r_id = idx + 2; // rId1 = master, slides start at rId2
        s.push_str(&format!("<p:sldId id=\"{}\" r:id=\"rId{}\"/>", slide.slide_id, r_id));
    }
    s.push_str("</p:sldIdLst>");
    // Layout size
    s.push_str(&format!(
        "<p:sldSz cx=\"{}\" cy=\"{}\"/>",
        pres.layout.width, pres.layout.height
    ));
    s.push_str(&format!(
        "<p:notesSz cx=\"{}\" cy=\"{}\"/>",
        pres.layout.height, pres.layout.width
    ));
    // Default text styles
    s.push_str("<p:defaultTextStyle>");
    for idy in 1..=9 {
        let mar_l = (idy - 1) * 457_200;
        s.push_str(&format!(
            "<a:lvl{idy}pPr marL=\"{mar_l}\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">\
<a:defRPr sz=\"1800\" kern=\"1200\">\
<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>\
<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\
</a:defRPr></a:lvl{idy}pPr>"
        ));
    }
    s.push_str("</p:defaultTextStyle>");

    // Sections (p14 extension)
    if !pres.sections.is_empty() {
        s.push_str("<p:extLst><p:ext uri=\"{521415D9-36F7-43E2-AB2F-B90AF26B5E84}\">");
        s.push_str("<p14:sectionLst xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\">");
        for (sec_idx, sec) in pres.sections.iter().enumerate() {
            // Generate a deterministic section GUID from index
            let guid = format!("{{{:08X}-0000-0000-0000-{:012X}}}", sec_idx + 1, sec_idx + 1);
            let sec_name = encode_xml_entities(&sec.name);
            s.push_str(&format!("<p14:section name=\"{sec_name}\" id=\"{guid}\">"));
            s.push_str("<p14:sldIdLst>");
            // Find slides that belong to this section (from start_slide to next section start - 1)
            let next_start = pres.sections.get(sec_idx + 1).map(|ns| ns.start_slide).unwrap_or(u32::MAX);
            for slide in &pres.slides {
                if slide.slide_num >= sec.start_slide && slide.slide_num < next_start {
                    s.push_str(&format!("<p14:sldId id=\"{}\"/>", slide.slide_id));
                }
            }
            s.push_str("</p14:sldIdLst>");
            s.push_str("</p14:section>");
        }
        s.push_str("</p14:sectionLst>");
        s.push_str("</p:ext>");
        s.push_str("<p:ext uri=\"{EFAFB233-063F-42B5-8137-9DF3F51BA10A}\">\
            <p15:sldGuideLst xmlns:p15=\"http://schemas.microsoft.com/office/powerpoint/2012/main\"/>\
            </p:ext>");
        s.push_str("</p:extLst>");
    }

    s.push_str("</p:presentation>");
    s
}

// ─────────────────────────────────────────────────────────────
// ppt/presProps.xml  (static)
// ─────────────────────────────────────────────────────────────

pub fn make_xml_pres_props() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
<p:presentationPr \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \
xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>"
}

// ─────────────────────────────────────────────────────────────
// ppt/tableStyles.xml  (static)
// ─────────────────────────────────────────────────────────────

pub fn make_xml_table_styles() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
<a:tblStyleLst \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
def=\"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}\"/>"
}

// ─────────────────────────────────────────────────────────────
// ppt/viewProps.xml  (static)
// ─────────────────────────────────────────────────────────────

pub fn make_xml_view_props() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
<p:viewPr \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \
xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\
<p:normalViewPr horzBarState=\"maximized\">\
<p:restoredLeft sz=\"15611\"/><p:restoredTop sz=\"94610\"/>\
</p:normalViewPr>\
<p:slideViewPr>\
<p:cSldViewPr snapToGrid=\"0\" snapToObjects=\"1\">\
<p:cViewPr varScale=\"1\">\
<p:scale><a:sx n=\"136\" d=\"100\"/><a:sy n=\"136\" d=\"100\"/></p:scale>\
<p:origin x=\"216\" y=\"312\"/>\
</p:cViewPr>\
<p:guideLst/>\
</p:cSldViewPr>\
</p:slideViewPr>\
<p:notesTextViewPr><p:cViewPr>\
<p:scale><a:sx n=\"1\" d=\"1\"/><a:sy n=\"1\" d=\"1\"/></p:scale>\
<p:origin x=\"0\" y=\"0\"/>\
</p:cViewPr></p:notesTextViewPr>\
<p:gridSpacing cx=\"76200\" cy=\"76200\"/>\
</p:viewPr>"
}

// ─────────────────────────────────────────────────────────────
// ppt/theme/theme1.xml
// ─────────────────────────────────────────────────────────────

pub fn make_xml_theme(head_font_face: Option<&str>, body_font_face: Option<&str>) -> String {
    let major = head_font_face
        .map(|f| format!("<a:latin typeface=\"{f}\"/>"))
        .unwrap_or_else(|| "<a:latin typeface=\"Calibri Light\" panose=\"020F0302020204030204\"/>".to_string());
    let minor = body_font_face
        .map(|f| format!("<a:latin typeface=\"{f}\"/>"))
        .unwrap_or_else(|| "<a:latin typeface=\"Calibri\" panose=\"020F0502020204030204\"/>".to_string());

    format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">\
<a:themeElements>\
<a:clrScheme name=\"Office\">\
<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>\
<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>\
<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2>\
<a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>\
<a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1>\
<a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>\
<a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3>\
<a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>\
<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5>\
<a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>\
<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>\
<a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>\
</a:clrScheme>\
<a:fontScheme name=\"Office\">\
<a:majorFont>{major}<a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>\
<a:minorFont>{minor}<a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont>\
</a:fontScheme>\
<a:fmtScheme name=\"Office\">\
<a:fillStyleLst>\
<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>\
<a:gradFill rotWithShape=\"1\"><a:gsLst>\
<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs>\
<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs>\
<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs>\
</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>\
<a:gradFill rotWithShape=\"1\"><a:gsLst>\
<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs>\
<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs>\
<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs>\
</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>\
</a:fillStyleLst>\
<a:lnStyleLst>\
<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>\
<a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>\
<a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>\
</a:lnStyleLst>\
<a:effectStyleLst>\
<a:effectStyle><a:effectLst/></a:effectStyle>\
<a:effectStyle><a:effectLst/></a:effectStyle>\
<a:effectStyle><a:effectLst>\
<a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\" rotWithShape=\"0\">\
<a:srgbClr val=\"000000\"><a:alpha val=\"63000\"/></a:srgbClr>\
</a:outerShdw></a:effectLst></a:effectStyle>\
</a:effectStyleLst>\
<a:bgFillStyleLst>\
<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>\
<a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill>\
<a:gradFill rotWithShape=\"1\"><a:gsLst>\
<a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs>\
<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs>\
<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs>\
</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>\
</a:bgFillStyleLst>\
</a:fmtScheme>\
</a:themeElements>\
<a:objectDefaults/><a:extraClrSchemeLst/>\
</a:theme>")
}

// ─────────────────────────────────────────────────────────────
// ppt/slideMasters/slideMaster1.xml
// ─────────────────────────────────────────────────────────────

/// Offset all image/media rIds in a vec of objects so they don't collide
/// with layout and theme rIds in the master rels file.
fn remap_object_rids(objects: &mut [crate::objects::SlideObject], offset: u32) {
    use crate::objects::SlideObject;
    for obj in objects.iter_mut() {
        match obj {
            SlideObject::Image(img) => { img.image_rid += offset; }
            SlideObject::Media(m) => {
                m.media_rid += offset;
                if let Some(ref mut pr) = m.poster_rid { *pr += offset; }
            }
            SlideObject::Group(g) => { remap_object_rids(&mut g.children, offset); }
            _ => {}
        }
    }
}

pub fn make_xml_master(pres: &Presentation) -> String {
    let layout_defs: String = pres.slide_layouts.iter().enumerate().map(|(idx, _layout)| {
        let id = crate::enums::LAYOUT_IDX_SERIES_BASE + idx as u64;
        format!("<p:sldLayoutId id=\"{id}\" r:id=\"rId{}\"/>", idx + 1)
    }).collect();

    // Master rId offset: layout rIds (1..N) + theme rId (N+1), so user content starts at N+2
    let rid_offset = (pres.slide_layouts.len() + 1) as u32;

    let mut s = String::new();
    s.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    s.push_str("<p:sldMaster \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \
xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">");

    // Background
    let master_name = pres.master.as_ref().map(|m| m.title.as_str()).unwrap_or("DEFAULT");
    s.push_str(&format!("<p:cSld name=\"{master_name}\">"));
    if let Some(ref m) = pres.master {
        if let Some(ref bg_image_rid) = m.background_image_rid {
            let actual_rid = bg_image_rid + rid_offset;
            s.push_str(&format!("<p:bg><p:bgPr><a:blipFill><a:blip r:embed=\"rId{actual_rid}\"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>"));
        } else if let Some(ref bg_color) = m.background_color {
            let c = bg_color.trim_start_matches('#').to_uppercase();
            if let Some(transparency) = m.background_transparency {
                let alpha = ((100.0 - transparency) * 1000.0) as u32;
                s.push_str(&format!("<p:bg><p:bgPr><a:solidFill><a:srgbClr val=\"{c}\"><a:alpha val=\"{alpha}\"/></a:srgbClr></a:solidFill><a:effectLst/></p:bgPr></p:bg>"));
            } else {
                s.push_str(&format!("<p:bg><p:bgPr><a:solidFill><a:srgbClr val=\"{c}\"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>"));
            }
        } else {
            s.push_str("<p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg>");
        }
    } else {
        s.push_str("<p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg>");
    }

    s.push_str("<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>");
    s.push_str("<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>");

    // Render master objects (remap rIds for images/media)
    if let Some(ref m) = pres.master {
        if !m.objects.is_empty() {
            let mut objects = m.objects.clone();
            remap_object_rids(&mut objects, rid_offset);
            s.push_str(&crate::xml::slide_xml::gen_xml_objects(&objects, pres));
        }
    }

    s.push_str("</p:spTree></p:cSld>");
    s.push_str("<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>");
    s.push_str(&format!("<p:sldLayoutIdLst>{layout_defs}</p:sldLayoutIdLst>"));
    s.push_str("<p:hf sldNum=\"0\" hdr=\"0\" ftr=\"0\" dt=\"0\"/>");
    s.push_str("<p:txStyles>\
<p:titleStyle>\
<a:lvl1pPr algn=\"ctr\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">\
<a:spcBef><a:spcPct val=\"0\"/></a:spcBef><a:buNone/>\
<a:defRPr sz=\"4400\" kern=\"1200\">\
<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>\
<a:latin typeface=\"+mj-lt\"/><a:ea typeface=\"+mj-ea\"/><a:cs typeface=\"+mj-cs\"/>\
</a:defRPr></a:lvl1pPr>\
</p:titleStyle>\
<p:bodyStyle>\
<a:lvl1pPr marL=\"342900\" indent=\"-342900\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">\
<a:spcBef><a:spcPct val=\"20000\"/></a:spcBef>\
<a:buFont typeface=\"Arial\" pitchFamily=\"34\" charset=\"0\"/><a:buChar char=\"&#x2022;\"/>\
<a:defRPr sz=\"3200\" kern=\"1200\">\
<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>\
<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\
</a:defRPr></a:lvl1pPr>\
</p:bodyStyle>\
<p:otherStyle>\
<a:defPPr><a:defRPr lang=\"en-US\"/></a:defPPr>\
<a:lvl1pPr marL=\"0\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">\
<a:defRPr sz=\"1800\" kern=\"1200\">\
<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>\
<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\
</a:defRPr></a:lvl1pPr>\
</p:otherStyle>\
</p:txStyles>");
    s.push_str("</p:sldMaster>");
    s
}

// ─────────────────────────────────────────────────────────────
// ppt/slideLayouts/slideLayoutN.xml
// ─────────────────────────────────────────────────────────────

pub fn make_xml_layout(layout: &SlideLayout) -> String {
    let name = &layout.name;
    // Note: slide dimensions are presentation-level (p:sldSz in presentation.xml).
    // p:sldLayout does NOT support cx/cy attributes — layout.width/height are stored
    // for potential future use (e.g. informing presentation.xml) but not emitted here.
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
<p:sldLayout \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \
xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" \
preserve=\"1\">\
<p:cSld name=\"{name}\">\
<p:spTree>\
<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>\
<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>\
<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>\
</p:spTree>\
</p:cSld>\
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>\
</p:sldLayout>"
    )
}

// ─────────────────────────────────────────────────────────────
// ppt/slides/slideN.xml  (wrapper — content from slide_xml)
// ─────────────────────────────────────────────────────────────

pub fn make_xml_slide(slide: &Slide, pres: &Presentation) -> String {
    let hidden_attr = if slide.hidden { " show=\"0\"" } else { "" };
    let content    = crate::xml::slide_xml::slide_object_to_xml(slide, pres);
    let timing     = crate::xml::slide_xml::gen_xml_timing(&slide.objects);
    let transition = crate::xml::slide_xml::gen_xml_transition(slide);

    // Check if any text run on this slide has equation OMML
    let has_equations = slide.objects.iter().any(|obj| {
        if let crate::objects::SlideObject::Text(t) = obj {
            t.text.iter().any(|r| r.equation_omml.is_some())
        } else {
            false
        }
    });
    let math_ns = if has_equations {
        " xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" \
xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" \
xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""
    } else {
        ""
    };

    // p:sld child order: p:cSld → p:clrMapOvr → p:transition → p:timing
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{CRLF}\
<p:sld \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \
xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"{math_ns}{hidden_attr}>\
{content}\
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>\
{transition}\
{timing}\
</p:sld>"
    )
}

// ─────────────────────────────────────────────────────────────
// ppt/notesSlides/notesSlideN.xml
// ─────────────────────────────────────────────────────────────

pub fn make_xml_notes_slide(slide: &Slide) -> String {
    let raw = slide.notes.as_deref().unwrap_or("");
    let notes_text = crate::utils::encode_xml_entities(raw);
    // Only emit a text run if there is actual note content
    let body_para = if raw.is_empty() {
        "<a:p><a:endParaRPr lang=\"en-US\" dirty=\"0\"/></a:p>".to_string()
    } else {
        format!("<a:p><a:r><a:rPr lang=\"en-US\" dirty=\"0\"/><a:t>{notes_text}</a:t></a:r><a:endParaRPr lang=\"en-US\" dirty=\"0\"/></a:p>")
    };
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>{CRLF}\
<p:notes \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \
xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\
<p:cSld><p:spTree>\
<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>\
<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>\
<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>\
<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Slide Image Placeholder 1\"/>\
<p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\"/></p:cNvSpPr>\
<p:nvPr><p:ph type=\"sldImg\"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>\
<p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Notes Placeholder 2\"/>\
<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>\
<p:nvPr><p:ph type=\"body\" idx=\"1\"/></p:nvPr></p:nvSpPr>\
<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{body_para}</p:txBody></p:sp>\
<p:sp><p:nvSpPr><p:cNvPr id=\"4\" name=\"Slide Number Placeholder 3\"/>\
<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>\
<p:nvPr><p:ph type=\"sldNum\" idx=\"5\" sz=\"quarter\"/></p:nvPr></p:nvSpPr>\
<p:spPr/></p:sp>\
</p:spTree>\
</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>",
        body_para = body_para
    )
}

// ─────────────────────────────────────────────────────────────
// ppt/notesMasters/notesMaster1.xml  (static)
// ─────────────────────────────────────────────────────────────

pub fn make_xml_notes_master() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
<p:notesMaster \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \
xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\
<p:cSld>\
<p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg>\
<p:spTree>\
<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>\
<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>\
<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>\
</p:spTree>\
</p:cSld>\
<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" \
accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" \
accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" \
hlink=\"hlink\" folHlink=\"folHlink\"/>\
<p:notesStyle>\
<a:lvl1pPr marL=\"0\" algn=\"l\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\">\
<a:defRPr sz=\"1200\" kern=\"1200\">\
<a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill>\
<a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/>\
</a:defRPr></a:lvl1pPr>\
</p:notesStyle>\
</p:notesMaster>"
}
