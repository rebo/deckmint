use std::io::Write;
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

use crate::error::PptxError;
use crate::presentation::Presentation;
use crate::xml::{chart_xml, pres_xml, rels_xml};

/// Assembles the complete PPTX ZIP archive into an in-memory buffer.
/// Follows the OOXML packaging convention for `.pptx` files.
pub fn assemble_pptx(pres: &Presentation) -> Result<Vec<u8>, PptxError> {
    let buf = std::io::Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(buf);

    let opts = SimpleFileOptions::default();
    let opts_deflate = SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // Helper: write a string file
    macro_rules! add_file {
        ($path:expr, $content:expr) => {{
            zip.start_file($path, opts_deflate).map_err(PptxError::Zip)?;
            zip.write_all($content.as_bytes()).map_err(PptxError::Io)?;
        }};
    }
    macro_rules! add_file_str {
        ($path:expr, $content:expr) => {{
            zip.start_file($path, opts_deflate).map_err(PptxError::Zip)?;
            zip.write_all($content).map_err(PptxError::Io)?;
        }};
    }

    // ── Root ──────────────────────────────────────────────
    add_file!("[Content_Types].xml", pres_xml::make_xml_content_types(pres));
    add_file!("_rels/.rels", pres_xml::make_xml_root_rels());

    // ── docProps ──────────────────────────────────────────
    add_file!("docProps/app.xml", pres_xml::make_xml_app(&pres.slides, &pres.company));
    add_file!("docProps/core.xml", pres_xml::make_xml_core(&pres.title, &pres.subject, &pres.author, &pres.revision));

    // ── ppt/ ──────────────────────────────────────────────
    add_file!("ppt/_rels/presentation.xml.rels", rels_xml::make_xml_presentation_rels(pres.slides.len()));
    add_file!("ppt/theme/theme1.xml", pres_xml::make_xml_theme(
        pres.theme.as_ref().and_then(|t| t.head_font_face.as_deref()),
        pres.theme.as_ref().and_then(|t| t.body_font_face.as_deref()),
    ));
    // Notes master requires its own dedicated theme file (theme2.xml) to avoid repair prompts
    add_file!("ppt/theme/theme2.xml", pres_xml::make_xml_theme(
        pres.theme.as_ref().and_then(|t| t.head_font_face.as_deref()),
        pres.theme.as_ref().and_then(|t| t.body_font_face.as_deref()),
    ));
    add_file!("ppt/presentation.xml", pres_xml::make_xml_presentation(pres));
    add_file_str!("ppt/presProps.xml", pres_xml::make_xml_pres_props().as_bytes());
    add_file_str!("ppt/tableStyles.xml", pres_xml::make_xml_table_styles().as_bytes());
    add_file_str!("ppt/viewProps.xml", pres_xml::make_xml_view_props().as_bytes());

    // ── Slide layouts ────────────────────────────────────
    for (idx, layout) in pres.slide_layouts.iter().enumerate() {
        let n = idx + 1;
        add_file!(
            &format!("ppt/slideLayouts/slideLayout{n}.xml"),
            pres_xml::make_xml_layout(layout)
        );
        add_file_str!(
            &format!("ppt/slideLayouts/_rels/slideLayout{n}.xml.rels"),
            rels_xml::make_xml_slide_layout_rel().as_bytes()
        );
    }

    // ── Slide master ─────────────────────────────────────
    add_file!("ppt/slideMasters/slideMaster1.xml", pres_xml::make_xml_master(pres));
    add_file!("ppt/slideMasters/_rels/slideMaster1.xml.rels", rels_xml::make_xml_master_rel(pres.slide_layouts.len()));

    // ── Notes master ─────────────────────────────────────
    add_file_str!("ppt/notesMasters/notesMaster1.xml", pres_xml::make_xml_notes_master().as_bytes());
    add_file_str!("ppt/notesMasters/_rels/notesMaster1.xml.rels", rels_xml::make_xml_notes_master_rel().as_bytes());

    // Pre-compute global sequential media targets with deduplication.
    // Media items with identical bytes reuse the same target path.
    let mut global_media_idx = 0u32;
    let mut flat_media: Vec<&crate::objects::SlideRelMedia> = Vec::new();
    for slide in &pres.slides {
        for media in &slide.rels_media {
            flat_media.push(media);
        }
    }
    // Build dedup: for each media, find if an earlier media has identical data
    let mut media_target_for: Vec<String> = Vec::with_capacity(flat_media.len());
    for (i, media) in flat_media.iter().enumerate() {
        let mut found_dup = None;
        for (j, earlier) in flat_media[..i].iter().enumerate() {
            if earlier.extn == media.extn && earlier.data == media.data {
                found_dup = Some(j);
                break;
            }
        }
        if let Some(j) = found_dup {
            media_target_for.push(media_target_for[j].clone());
        } else {
            global_media_idx += 1;
            let prefix = if media.rel_type == "video" || media.rel_type == "audio" { "media" } else { "image" };
            media_target_for.push(format!("../media/{prefix}{global_media_idx}.{}", media.extn));
        }
    }
    // Split back into per-slide targets
    let mut flat_idx = 0;
    let slide_media_targets: Vec<Vec<String>> = pres.slides.iter().map(|slide| {
        let targets: Vec<String> = slide.rels_media.iter().map(|_| {
            let t = media_target_for[flat_idx].clone();
            flat_idx += 1;
            t
        }).collect();
        targets
    }).collect();

    // Pre-compute global sequential chart targets (chart1, chart2, … across all slides)
    let mut global_chart_idx = 0u32;
    let slide_chart_rels: Vec<Vec<(u32, String)>> = pres.slides.iter().map(|slide| {
        slide.charts.iter().map(|chart| {
            global_chart_idx += 1;
            (chart.chart_rid, format!("../charts/chart{}.xml", global_chart_idx))
        }).collect()
    }).collect();

    // ── Slides ───────────────────────────────────────────
    for (idx, slide) in pres.slides.iter().enumerate() {
        let n = idx + 1;
        let layout_idx = 1; // All slides use layout 1 (DEFAULT)

        add_file!(&format!("ppt/slides/slide{n}.xml"), pres_xml::make_xml_slide(slide, pres));
        add_file!(
            &format!("ppt/slides/_rels/slide{n}.xml.rels"),
            rels_xml::make_xml_slide_rel(n, layout_idx, &slide.rels, &slide.rels_media, &slide_media_targets[idx], &slide_chart_rels[idx])
        );
        add_file!(&format!("ppt/notesSlides/notesSlide{n}.xml"), pres_xml::make_xml_notes_slide(slide));
        add_file!(&format!("ppt/notesSlides/_rels/notesSlide{n}.xml.rels"), rels_xml::make_xml_notes_slide_rel(n));
    }

    // ── Charts ──────────────────────────────────────────
    let mut global_chart_idx = 0u32;
    for slide in &pres.slides {
        for chart in &slide.charts {
            global_chart_idx += 1;
            add_file!(
                &format!("ppt/charts/chart{global_chart_idx}.xml"),
                chart_xml::gen_xml_chart(chart)
            );
            add_file_str!(
                &format!("ppt/charts/_rels/chart{global_chart_idx}.xml.rels"),
                chart_xml::gen_xml_chart_rels().as_bytes()
            );
        }
    }

    // ── Media files (deduplicated) ─────────────────────
    let mut written_paths: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut flat_idx2 = 0usize;
    for slide in &pres.slides {
        for media in &slide.rels_media {
            let target = &media_target_for[flat_idx2];
            flat_idx2 += 1;
            // Strip leading "../" to get the ZIP path
            let zip_path = format!("ppt/{}", target.trim_start_matches("../"));
            if written_paths.insert(zip_path.clone()) {
                zip.start_file(&zip_path, opts).map_err(PptxError::Zip)?;
                zip.write_all(&media.data).map_err(PptxError::Io)?;
            }
        }
    }

    let result = zip.finish().map_err(PptxError::Zip)?;
    Ok(result.into_inner())
}
