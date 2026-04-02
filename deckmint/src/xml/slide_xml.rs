use crate::enums::EMU;
use crate::objects::{SlideObject, TextObject, ShapeObject, ImageObject, ConnectorObject, MediaObject, GroupObject};
use crate::objects::text::{TextOptions, TextRun, BulletType, TextFit};
use crate::objects::image::ImageSizingType;
use crate::presentation::Presentation;
use crate::slide::Slide;
use crate::types::{AnimationEffectType, CheckerboardDir, Direction, FillType, PatternFill, ShadowType, ShapeVariant, SplitOrientation, StripDir, TransitionDir};
use crate::utils::{
    convert_rotation_degrees, create_glow_element, encode_xml_entities,
    gen_xml_color_selection_str, inch_to_emu, val_to_pts,
};
use crate::xml::table_xml::gen_xml_table;

// ─────────────────────────────────────────────────────────────
// Render a list of SlideObjects into <p:sp>/<p:pic> fragments
// Used by both slides and slide master
// ─────────────────────────────────────────────────────────────

pub fn gen_xml_objects(objects: &[SlideObject], pres: &Presentation) -> String {
    let mut s = String::new();
    let mut id_counter: usize = 2;
    for obj in objects {
        render_slide_object(obj, &mut id_counter, pres, &mut s);
    }
    s
}

/// Generate a fallback text object with equations stripped (for mc:Fallback).
fn gen_xml_text_object_fallback(t: &TextObject, obj_id: usize, pres: &Presentation) -> String {
    let (x, y, cx, cy) = resolve_position(&t.options.position, pres);
    let obj_name = &t.object_name;
    let mut s = String::new();
    s.push_str("<p:sp>");
    s.push_str(&format!(
        "<p:nvSpPr>\
<p:cNvPr id=\"{obj_id}\" name=\"{obj_name}\"/>\
<p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr>\
<p:nvPr/>\
</p:nvSpPr>"
    ));
    s.push_str(&format!(
        "<p:spPr>\
<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>\
<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>\
<a:noFill/>\
</p:spPr>"
    ));
    // Emit only plain text runs (skip equations)
    s.push_str("<p:txBody><a:bodyPr/><a:lstStyle/>");
    s.push_str("<a:p>");
    for run in &t.text {
        if run.equation_omml.is_none() && !run.text.is_empty() {
            s.push_str(&format!("<a:r><a:rPr lang=\"en-US\"/><a:t>{}</a:t></a:r>",
                encode_xml_entities(&run.text)));
        }
    }
    s.push_str("<a:endParaRPr lang=\"en-US\"/>");
    s.push_str("</a:p></p:txBody></p:sp>");
    s
}

// ─────────────────────────────────────────────────────────────
// Top-level: <p:cSld> content for a slide
// ─────────────────────────────────────────────────────────────

pub fn slide_object_to_xml(slide: &Slide, pres: &Presentation) -> String {
    let mut s = String::new();

    // Open <p:cSld>
    s.push_str("<p:cSld>");

    // Background
    if let Some(ref bg) = slide.background {
        if let Some(rid) = bg.image_rid {
            s.push_str("<p:bg><p:bgPr>");
            s.push_str(&format!(
                "<a:blipFill dpi=\"0\" rotWithShape=\"1\">\
<a:blip r:embed=\"rId{rid}\"/>\
<a:stretch><a:fillRect/></a:stretch>\
</a:blipFill><a:effectLst/>"
            ));
            s.push_str("</p:bgPr></p:bg>");
        } else if let Some(ref color) = bg.color {
            s.push_str("<p:bg><p:bgPr>");
            s.push_str(&gen_xml_color_selection_str(color, bg.transparency));
            s.push_str("<a:effectLst/>");
            s.push_str("</p:bgPr></p:bg>");
        }
    } else {
        // Default white background for master reference
        s.push_str("<p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg>");
    }

    // Shape tree — cNvPr ids must be unique within this slide; id=1 is reserved for the group.
    s.push_str("<p:spTree>");
    s.push_str("<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>");
    s.push_str("<p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>");

    // Render each object — use a mutable counter for IDs (groups consume multiple IDs)
    let mut id_counter: usize = 2; // id=1 reserved for spTree root group
    for obj in &slide.objects {
        render_slide_object(obj, &mut id_counter, pres, &mut s);
    }

    // Slide number placeholder
    if let Some(ref sn) = slide.slide_number {
        let sn_id = id_counter;
        id_counter += 1;
        s.push_str(&gen_xml_slide_number(sn, sn_id, pres));
    }

    // Charts (graphicFrames — rendered inside spTree but after regular shapes)
    for chart in &slide.charts {
        let obj_id = id_counter;
        id_counter += 1;
        s.push_str(&crate::xml::chart_xml::gen_xml_chart_frame(chart, obj_id, pres));
    }
    let _ = id_counter; // suppress unused warning

    s.push_str("</p:spTree>");
    s.push_str("</p:cSld>");
    s
}

/// Render a single SlideObject to XML, advancing the mutable ID counter.
/// Groups recursively render their children.
fn render_slide_object(obj: &SlideObject, id_counter: &mut usize, pres: &Presentation, s: &mut String) {
    let obj_id = *id_counter;
    *id_counter += 1;
    match obj {
        SlideObject::Text(t) => {
            let sp_xml = gen_xml_text_object(t, obj_id, pres);
            let has_eq = t.text.iter().any(|r| r.equation_omml.is_some());
            if has_eq {
                s.push_str("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">");
                s.push_str("<mc:Choice Requires=\"a14\">");
                s.push_str(&sp_xml);
                s.push_str("</mc:Choice>");
                s.push_str("<mc:Fallback>");
                let fallback_id = obj_id + 10000;
                s.push_str(&gen_xml_text_object_fallback(t, fallback_id, pres));
                s.push_str("</mc:Fallback>");
                s.push_str("</mc:AlternateContent>");
            } else {
                s.push_str(&sp_xml);
            }
        }
        SlideObject::Shape(sh) => s.push_str(&gen_xml_shape_object(sh, obj_id, pres)),
        SlideObject::Image(img) => s.push_str(&gen_xml_image_object(img, obj_id, pres)),
        SlideObject::Table(tbl) => s.push_str(&gen_xml_table(tbl, obj_id, pres)),
        SlideObject::Connector(c) => s.push_str(&gen_xml_connector_object(c, obj_id, pres)),
        SlideObject::Media(m) => s.push_str(&gen_xml_media_object(m, obj_id, pres)),
        SlideObject::Group(g) => {
            s.push_str(&gen_xml_group_object(g, obj_id, id_counter, pres));
        }
    }
}

// ─────────────────────────────────────────────────────────────
// Text object → <p:sp>
// ─────────────────────────────────────────────────────────────

fn gen_xml_text_object(t: &TextObject, obj_id: usize, pres: &Presentation) -> String {
    let (x, y, cx, cy) = resolve_position(&t.options.position, pres);
    let location_attr = build_location_attr(t.options.rotate, t.options.flip_h, t.options.flip_v);
    let obj_name = &t.object_name;

    let mut s = String::new();
    s.push_str("<p:sp>");
    s.push_str(&format!(
        "<p:nvSpPr>\
<p:cNvPr id=\"{obj_id}\" name=\"{obj_name}\"/>\
<p:cNvSpPr txBox=\"1\"><a:spLocks noGrp=\"1\"/></p:cNvSpPr>\
<p:nvPr/>\
</p:nvSpPr>"
    ));
    s.push_str("<p:spPr>");
    s.push_str(&format!(
        "<a:xfrm{location_attr}><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>"
    ));
    s.push_str("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");

    // Fill
    if let Some(ref grad) = t.options.gradient_fill {
        s.push_str(&crate::utils::gen_xml_grad_fill(grad));
    } else if let Some(ref fill_color) = t.options.fill {
        s.push_str(&gen_xml_color_selection_str(fill_color, None));
    } else {
        s.push_str("<a:noFill/>");
    }

    // Border line
    if let Some(ref line) = t.options.line {
        s.push_str(&gen_xml_shape_line(line));
    }

    // Shadow
    if let Some(ref shadow) = t.options.shadow {
        if shadow.shadow_type != ShadowType::None {
            s.push_str(&gen_xml_shadow(shadow));
        }
    }

    s.push_str("</p:spPr>");

    // Text body
    s.push_str(&gen_xml_text_body_from_runs(&t.text, &t.options, false));
    s.push_str("</p:sp>");
    s
}

// ─────────────────────────────────────────────────────────────
// Shape object → <p:sp>
// ─────────────────────────────────────────────────────────────

fn gen_xml_shape_object(sh: &ShapeObject, obj_id: usize, pres: &Presentation) -> String {
    let (x, y, cx, cy) = resolve_position(&sh.options.position, pres);
    let location_attr = build_location_attr(
        sh.options.rotate,
        sh.options.flip_h,
        sh.options.flip_v,
    );
    let obj_name = &sh.object_name;
    let shape_prst = sh.shape_type.as_str();

    let mut s = String::new();
    s.push_str("<p:sp>");

    // cNvPr — include alt text and/or hlinkClick/hlinkHover children
    let alt_attr = sh.options.alt_text.as_deref()
        .map(|a| format!(" descr=\"{}\"", encode_xml_entities(a)))
        .unwrap_or_default();
    let cnv_pr = {
        let mut children = String::new();
        if let Some(ref hl) = sh.options.hyperlink {
            children.push_str(&gen_xml_hlink(hl, "hlinkClick"));
        }
        if let Some(ref hl) = sh.options.hover {
            children.push_str(&gen_xml_hlink(hl, "hlinkHover"));
        }
        if children.is_empty() {
            format!("<p:cNvPr id=\"{obj_id}\" name=\"{obj_name}\"{alt_attr}/>")
        } else {
            format!("<p:cNvPr id=\"{obj_id}\" name=\"{obj_name}\"{alt_attr}>{children}</p:cNvPr>")
        }
    };

    s.push_str(&format!(
        "<p:nvSpPr>\
{cnv_pr}\
<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>\
<p:nvPr/>\
</p:nvSpPr>"
    ));
    s.push_str("<p:spPr>");
    s.push_str(&format!(
        "<a:xfrm{location_attr}><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>"
    ));

    // Custom or preset geometry
    if let Some(ref pts) = sh.options.custom_geometry {
        s.push_str(&gen_xml_custom_geom(pts, cx, cy));
    } else {
    s.push_str(&format!("<a:prstGeom prst=\"{shape_prst}\">"));
    if let Some(r) = sh.options.rect_radius {
        let adj = ((r * EMU as f64 * 100_000.0) / cx.max(cy) as f64).round() as i64;
        s.push_str(&format!("<a:avLst><a:gd name=\"adj\" fmla=\"val {adj}\"/></a:avLst>"));
    } else if let Some([start, swing]) = sh.options.angle_range {
        use crate::enums::ShapeType as ST;
        let adj1 = (start * 60_000.0).round() as i64;
        let adj2 = (swing * 60_000.0).round() as i64;
        match sh.shape_type {
            ST::BlockArc => {
                // adj1 = start, adj2 = swing, adj3 = inner radius (0–100000)
                let inner = ((1.0 - sh.options.arc_thickness.unwrap_or(0.5).clamp(0.0, 1.0)) * 100_000.0).round() as i64;
                s.push_str(&format!(
                    "<a:avLst>\
<a:gd name=\"adj1\" fmla=\"val {adj1}\"/>\
<a:gd name=\"adj2\" fmla=\"val {adj2}\"/>\
<a:gd name=\"adj3\" fmla=\"val {inner}\"/>\
</a:avLst>"
                ));
            }
            _ => {
                // pie / arc / pieWedge — adj1 = start, adj2 = swing
                s.push_str(&format!(
                    "<a:avLst>\
<a:gd name=\"adj1\" fmla=\"val {adj1}\"/>\
<a:gd name=\"adj2\" fmla=\"val {adj2}\"/>\
</a:avLst>"
                ));
            }
        }
    } else {
        s.push_str("<a:avLst/>");
    }
    s.push_str("</a:prstGeom>");
    } // end custom_geometry else

    // Fill
    match &sh.options.fill {
        Some(fill) => {
            match fill.fill_type {
                FillType::None => s.push_str("<a:noFill/>"),
                FillType::Solid => {
                    if let Some(ref color) = fill.color {
                        s.push_str(&crate::utils::gen_xml_color_selection(color, fill.transparency));
                    } else {
                        s.push_str("<a:noFill/>");
                    }
                }
                FillType::Gradient => {
                    if let Some(ref grad) = fill.gradient {
                        s.push_str(&crate::utils::gen_xml_grad_fill(grad));
                    } else {
                        s.push_str("<a:noFill/>");
                    }
                }
                FillType::Pattern => {
                    if let Some(ref pf) = fill.pattern {
                        s.push_str(&gen_xml_pattern_fill(pf));
                    } else {
                        s.push_str("<a:noFill/>");
                    }
                }
            }
        }
        None => s.push_str("<a:noFill/>"),
    }

    // Line
    if let Some(ref line) = sh.options.line {
        s.push_str(&gen_xml_shape_line(line));
    }

    // Shadow
    if let Some(ref shadow) = sh.options.shadow {
        if shadow.shadow_type != ShadowType::None {
            s.push_str(&gen_xml_shadow(shadow));
        }
    }

    // 3D scene (camera + light rig)
    if let Some(ref scene) = sh.options.scene_3d {
        s.push_str(&gen_xml_scene3d(scene));
    }

    // 3D shape (bevel, extrusion, material)
    if let Some(ref sp3d) = sh.options.shape_3d {
        s.push_str(&gen_xml_sp3d(sp3d));
    }

    s.push_str("</p:spPr>");

    // Optional text body
    if let Some(ref text) = sh.text {
        s.push_str(&gen_xml_text_body_from_runs(&text.text, &text.options, false));
    } else {
        // Empty text body required by OOXML spec for shapes
        s.push_str("<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\" dirty=\"0\"/></a:p></p:txBody>");
    }

    s.push_str("</p:sp>");
    s
}

// ─────────────────────────────────────────────────────────────
// Image object → <p:pic>
// ─────────────────────────────────────────────────────────────

fn gen_xml_image_object(img: &ImageObject, obj_id: usize, pres: &Presentation) -> String {
    let (x, y, cx, cy) = resolve_position(&img.options.position, pres);
    let mut img_w = cx;
    let mut img_h = cy;
    let obj_name = &img.object_name;
    let alt_text = encode_xml_entities(img.options.alt_text.as_deref().unwrap_or(""));
    let rid = img.image_rid;

    let location_attr = build_location_attr(img.options.rotate, img.options.flip_h, img.options.flip_v);

    let mut s = String::new();
    s.push_str("<p:pic>");
    s.push_str("<p:nvPicPr>");

    // cNvPr — include hlinkClick/hlinkHover children
    {
        let mut children = String::new();
        if let Some(ref hl) = img.options.hyperlink {
            children.push_str(&gen_xml_hlink(hl, "hlinkClick"));
        }
        if let Some(ref hl) = img.options.hover {
            children.push_str(&gen_xml_hlink(hl, "hlinkHover"));
        }
        if children.is_empty() {
            s.push_str(&format!("<p:cNvPr id=\"{obj_id}\" name=\"{obj_name}\" descr=\"{alt_text}\"/>"));
        } else {
            s.push_str(&format!("<p:cNvPr id=\"{obj_id}\" name=\"{obj_name}\" descr=\"{alt_text}\">{children}</p:cNvPr>"));
        }
    }
    s.push_str("<p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr>");
    s.push_str("<p:nvPr/>");
    s.push_str("</p:nvPicPr>");

    // blipFill
    s.push_str("<p:blipFill>");
    if img.is_svg {
        s.push_str(&format!("<a:blip r:embed=\"rId{}\">", rid - 1));
        if let Some(t) = img.options.transparency {
            s.push_str(&format!("<a:alphaModFix amt=\"{}\"/>", ((100.0 - t) * 1000.0).round() as i64));
        }
        s.push_str("<a:extLst><a:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\">");
        s.push_str(&format!("<asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"rId{rid}\"/>"));
        s.push_str("</a:ext></a:extLst>");
        s.push_str("</a:blip>");
    } else {
        s.push_str(&format!("<a:blip r:embed=\"rId{rid}\">"));
        if let Some(t) = img.options.transparency {
            s.push_str(&format!("<a:alphaModFix amt=\"{}\"/>", ((100.0 - t) * 1000.0).round() as i64));
        }
        if let Some(ref adj) = img.options.color_adjust {
            if adj.brightness.is_some() || adj.contrast.is_some() {
                let bright = adj.brightness.map(|b| (b * 1000.0).round() as i64).unwrap_or(0);
                let contrast = adj.contrast.map(|c| (c * 1000.0).round() as i64).unwrap_or(0);
                s.push_str(&format!("<a:lum bright=\"{bright}\" contrast=\"{contrast}\"/>"));
            }
            if adj.grayscale {
                s.push_str("<a:grayscl/>");
            }
        }
        s.push_str("</a:blip>");
    }

    // Sizing
    if let Some(ref sizing) = img.options.sizing {
        let box_w = sizing.w.map(inch_to_emu).unwrap_or(cx);
        let box_h = sizing.h.map(inch_to_emu).unwrap_or(cy);
        let box_x = sizing.x.map(inch_to_emu).unwrap_or(0);
        let box_y = sizing.y.map(inch_to_emu).unwrap_or(0);
        let img_size = (cx, cy);
        let box_dim = (box_w, box_h, box_x, box_y);
        s.push_str(&image_sizing_xml(&sizing.sizing_type, img_size, box_dim));
        img_w = box_w;
        img_h = box_h;
    } else if let Some(ref crop) = img.options.crop {
        // Direct LTRB crop without sizing mode
        let l = (crop[0] * 100_000.0).round() as i64;
        let t = (crop[1] * 100_000.0).round() as i64;
        let r = (crop[2] * 100_000.0).round() as i64;
        let b = (crop[3] * 100_000.0).round() as i64;
        s.push_str(&format!("<a:srcRect l=\"{l}\" t=\"{t}\" r=\"{r}\" b=\"{b}\"/><a:stretch><a:fillRect/></a:stretch>"));
    } else {
        s.push_str("<a:stretch><a:fillRect/></a:stretch>");
    }
    s.push_str("</p:blipFill>");

    // spPr
    s.push_str("<p:spPr>");
    s.push_str(&format!(
        "<a:xfrm{location_attr}><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{img_w}\" cy=\"{img_h}\"/></a:xfrm>"
    ));
    let geom = if img.options.rounding { "ellipse" } else { "rect" };
    s.push_str(&format!("<a:prstGeom prst=\"{geom}\"><a:avLst/></a:prstGeom>"));

    if let Some(ref shadow) = img.options.shadow {
        if shadow.shadow_type != ShadowType::None {
            s.push_str(&gen_xml_shadow(shadow));
        }
    }

    s.push_str("</p:spPr>");
    s.push_str("</p:pic>");
    s
}

// ─────────────────────────────────────────────────────────────
// Image sizing XML helpers (cover / contain / crop)
// ─────────────────────────────────────────────────────────────

fn image_sizing_xml(
    sizing_type: &ImageSizingType,
    img_size: (i64, i64),
    box_dim: (i64, i64, i64, i64),
) -> String {
    let (img_w, img_h) = (img_size.0 as f64, img_size.1 as f64);
    let (box_w, box_h, box_x, box_y) = (box_dim.0 as f64, box_dim.1 as f64, box_dim.2 as f64, box_dim.3 as f64);

    match sizing_type {
        ImageSizingType::Cover => {
            let img_ratio = img_h / img_w;
            let box_ratio = box_h / box_w;
            let is_box_based = box_ratio > img_ratio;
            let width = if is_box_based { box_h / img_ratio } else { box_w };
            let height = if is_box_based { box_h } else { box_w * img_ratio };
            let hz = ((0.5 * (1.0 - box_w / width)) * 1e5).round() as i64;
            let vz = ((0.5 * (1.0 - box_h / height)) * 1e5).round() as i64;
            format!("<a:srcRect l=\"{hz}\" r=\"{hz}\" t=\"{vz}\" b=\"{vz}\"/><a:stretch/>")
        }
        ImageSizingType::Contain => {
            let img_ratio = img_h / img_w;
            let box_ratio = box_h / box_w;
            let width_based = box_ratio > img_ratio;
            let width = if width_based { box_w } else { box_h / img_ratio };
            let height = if width_based { box_w * img_ratio } else { box_h };
            let hz = ((0.5 * (1.0 - box_w / width)) * 1e5).round() as i64;
            let vz = ((0.5 * (1.0 - box_h / height)) * 1e5).round() as i64;
            format!("<a:srcRect l=\"{hz}\" r=\"{hz}\" t=\"{vz}\" b=\"{vz}\"/><a:stretch/>")
        }
        ImageSizingType::Crop => {
            let l = box_x;
            let r = img_w - (box_x + box_w);
            let t = box_y;
            let b = img_h - (box_y + box_h);
            let lp = ((l / img_w) * 1e5).round() as i64;
            let rp = ((r / img_w) * 1e5).round() as i64;
            let tp = ((t / img_h) * 1e5).round() as i64;
            let bp = ((b / img_h) * 1e5).round() as i64;
            format!("<a:srcRect l=\"{lp}\" r=\"{rp}\" t=\"{tp}\" b=\"{bp}\"/><a:stretch/>")
        }
    }
}

// ─────────────────────────────────────────────────────────────
// Group object → <p:grpSp>
// ─────────────────────────────────────────────────────────────

fn gen_xml_group_object(g: &GroupObject, obj_id: usize, id_counter: &mut usize, pres: &Presentation) -> String {
    let (x, y, cx, cy) = resolve_position(&g.position, pres);
    let obj_name = &g.object_name;
    let (ch_off_x, ch_off_y) = g.child_offset;
    let (ch_ext_cx, ch_ext_cy) = g.child_extent;

    let mut s = String::new();
    s.push_str("<p:grpSp>");

    // Non-visual properties
    s.push_str(&format!(
        "<p:nvGrpSpPr>\
<p:cNvPr id=\"{obj_id}\" name=\"{obj_name}\"/>\
<p:cNvGrpSpPr/>\
<p:nvPr/>\
</p:nvGrpSpPr>"
    ));

    // Group shape properties with child coordinate space
    s.push_str(&format!(
        "<p:grpSpPr><a:xfrm>\
<a:off x=\"{x}\" y=\"{y}\"/>\
<a:ext cx=\"{cx}\" cy=\"{cy}\"/>\
<a:chOff x=\"{ch_off_x}\" y=\"{ch_off_y}\"/>\
<a:chExt cx=\"{ch_ext_cx}\" cy=\"{ch_ext_cy}\"/>\
</a:xfrm></p:grpSpPr>"
    ));

    // Recursively render children
    for child in &g.children {
        render_slide_object(child, id_counter, pres, &mut s);
    }

    s.push_str("</p:grpSp>");
    s
}

// ─────────────────────────────────────────────────────────────
// Media object (video / audio) → <p:pic>
// ─────────────────────────────────────────────────────────────

fn gen_xml_media_object(m: &MediaObject, obj_id: usize, pres: &Presentation) -> String {
    let (x, y, cx, cy) = resolve_position(&m.options.position, pres);
    let obj_name = &m.object_name;
    let alt_text = m.options.alt_text.as_deref().unwrap_or("");
    let media_rid = m.media_rid;

    let mut s = String::new();
    s.push_str("<p:pic>");
    s.push_str("<p:nvPicPr>");
    s.push_str(&format!(
        "<p:cNvPr id=\"{obj_id}\" name=\"{obj_name}\" descr=\"{alt_text}\"/>"
    ));
    s.push_str("<p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr>");

    // nvPr contains the media file reference
    s.push_str("<p:nvPr>");
    match m.media_type {
        crate::objects::MediaType::Video => {
            s.push_str(&format!("<a:videoFile r:link=\"rId{media_rid}\"/>"));
        }
        crate::objects::MediaType::Audio => {
            s.push_str(&format!("<a:audioFile r:link=\"rId{media_rid}\"/>"));
        }
    }
    s.push_str("</p:nvPr>");
    s.push_str("</p:nvPicPr>");

    // blipFill — poster frame for video, placeholder for audio
    s.push_str("<p:blipFill>");
    if let Some(poster_rid) = m.poster_rid {
        s.push_str(&format!("<a:blip r:embed=\"rId{poster_rid}\"/>"));
    } else {
        // No poster frame — emit an empty blip (PowerPoint will show a generic icon)
        s.push_str("<a:blip/>");
    }
    s.push_str("<a:stretch><a:fillRect/></a:stretch>");
    s.push_str("</p:blipFill>");

    // spPr
    s.push_str("<p:spPr>");
    s.push_str(&format!(
        "<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>"
    ));
    s.push_str("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
    s.push_str("</p:spPr>");

    s.push_str("</p:pic>");
    s
}

// ─────────────────────────────────────────────────────────────
// Hyperlink XML helper (click and hover)
// ─────────────────────────────────────────────────────────────

/// Generate `<a:hlinkClick>` or `<a:hlinkHover>` XML for a hyperlink.
fn gen_xml_hlink(hl: &crate::types::HyperlinkProps, element: &str) -> String {
    let tooltip = hl.tooltip.as_deref().unwrap_or("");
    let tooltip_attr = if tooltip.is_empty() { String::new() } else { format!(" tooltip=\"{}\"", encode_xml_entities(tooltip)) };
    if let Some(ref nav) = hl.action {
        format!("<a:{element} r:id=\"\" action=\"{}\"{tooltip_attr}/>", nav.as_ppaction())
    } else if hl.r_id > 0 {
        if hl.slide.is_some() {
            format!("<a:{element} r:id=\"rId{}\" action=\"ppaction://hlinksldjump\"{tooltip_attr}/>", hl.r_id)
        } else {
            format!("<a:{element} r:id=\"rId{}\"{tooltip_attr}/>", hl.r_id)
        }
    } else {
        String::new()
    }
}

// ─────────────────────────────────────────────────────────────
// Text body generation
// ─────────────────────────────────────────────────────────────

pub fn gen_xml_text_body_from_runs(
    runs: &[TextRun],
    opts: &TextOptions,
    is_table_cell: bool,
) -> String {
    let tag = if is_table_cell { "a:txBody" } else { "p:txBody" };
    let mut s = String::new();
    s.push_str(&format!("<{tag}>"));

    // Body properties
    s.push_str(&gen_xml_body_properties(opts, is_table_cell));

    // List style
    s.push_str("<a:lstStyle/>");

    // Group runs into paragraphs
    let paragraphs = group_runs_into_paragraphs(runs);

    for para in &paragraphs {
        s.push_str("<a:p>");
        // Paragraph properties
        s.push_str(&gen_xml_para_props(opts, false));
        // Text runs
        for run in para {
            if run.text.is_empty() && run.field.is_none() && run.equation_omml.is_none() {
                s.push_str("<a:endParaRPr lang=\"en-US\" dirty=\"0\"/>");
            } else {
                // Soft line break before this run if requested
                if run.soft_break_before {
                    s.push_str("<a:br><a:rPr lang=\"en-US\" dirty=\"0\"/></a:br>");
                }
                s.push_str(&gen_xml_text_run(run, opts));
            }
        }
        s.push_str("<a:endParaRPr lang=\"en-US\" dirty=\"0\"/>");
        s.push_str("</a:p>");
    }

    if paragraphs.is_empty() {
        s.push_str("<a:p><a:endParaRPr lang=\"en-US\" dirty=\"0\"/></a:p>");
    }

    s.push_str(&format!("</{tag}>"));
    s
}

/// Group text runs into paragraphs based on break_line flags
fn group_runs_into_paragraphs(runs: &[TextRun]) -> Vec<Vec<&TextRun>> {
    let mut paragraphs: Vec<Vec<&TextRun>> = Vec::new();
    let mut current: Vec<&TextRun> = Vec::new();

    for run in runs {
        current.push(run);
        if run.break_line {
            paragraphs.push(current);
            current = Vec::new();
        }
    }
    if !current.is_empty() {
        paragraphs.push(current);
    }

    // Ensure at least one paragraph
    if paragraphs.is_empty() {
        paragraphs.push(Vec::new());
    }

    paragraphs
}

fn gen_xml_body_properties(opts: &TextOptions, is_table_cell: bool) -> String {
    if is_table_cell {
        return "<a:bodyPr/>".to_string();
    }

    let mut s = String::from("<a:bodyPr");

    let wrap = opts.wrap.unwrap_or(true);
    s.push_str(&format!(" wrap=\"{}\"", if wrap { "square" } else { "none" }));

    // Text direction (vert attribute)
    if let Some(ref vert) = opts.vert {
        s.push_str(&format!(" vert=\"{vert}\""));
    }

    // Margins (from body_prop or opts.margin)
    if let Some(ref bp) = opts.body_prop {
        if let Some(v) = bp.l_ins { s.push_str(&format!(" lIns=\"{v}\"")); }
        if let Some(v) = bp.t_ins { s.push_str(&format!(" tIns=\"{v}\"")); }
        if let Some(v) = bp.r_ins { s.push_str(&format!(" rIns=\"{v}\"")); }
        if let Some(v) = bp.b_ins { s.push_str(&format!(" bIns=\"{v}\"")); }
        if let Some(ref anchor) = bp.anchor { s.push_str(&format!(" anchor=\"{anchor}\"")); }
        if let Some(ref vert2) = bp.vert { s.push_str(&format!(" vert=\"{vert2}\"")); }
    } else if let Some(ref margin) = opts.margin {
        s.push_str(&format!(" lIns=\"{}\"", val_to_pts(margin.left())));
        s.push_str(&format!(" tIns=\"{}\"", val_to_pts(margin.top())));
        s.push_str(&format!(" rIns=\"{}\"", val_to_pts(margin.right())));
        s.push_str(&format!(" bIns=\"{}\"", val_to_pts(margin.bottom())));
        // Vertical alignment
        if let Some(ref valign) = opts.valign {
            s.push_str(&format!(" anchor=\"{}\"", valign.as_ooxml()));
        }
    }

    s.push_str(" rtlCol=\"0\"");

    // Multi-column layout
    if let Some(n) = opts.num_columns {
        if n >= 2 {
            s.push_str(&format!(" numCol=\"{n}\""));
            if let Some(gap) = opts.column_spacing {
                s.push_str(&format!(" spcCol=\"{}\"", inch_to_emu(gap)));
            }
        }
    }

    // Collect children
    let mut children = String::new();
    match opts.fit {
        Some(TextFit::None)   => children.push_str("<a:noAutofit/>"),
        Some(TextFit::Shrink) => children.push_str("<a:normAutofit/>"),
        Some(TextFit::Resize) => children.push_str("<a:spAutoFit/>"),
        _ => {}
    }
    if let Some(ref bp) = opts.body_prop {
        if bp.auto_fit { children.push_str("<a:spAutoFit/>"); }
    }

    if children.is_empty() {
        s.push_str("/>");
    } else {
        s.push('>');
        s.push_str(&children);
        s.push_str("</a:bodyPr>");
    }
    s
}

fn gen_xml_para_props(opts: &TextOptions, _is_default: bool) -> String {
    let mut s = String::new();
    let has_align = opts.align.is_some();
    let has_line = opts.line_spacing.is_some() || opts.line_spacing_multiple.is_some();
    let _has_bullet = opts.bullet.is_none();  // no bullet = add buNone
    let has_indent = opts.indent_level.map_or(false, |l| l > 0);
    let rtl = opts.rtl_mode;
    let has_tab_stops = opts.tab_stops.as_ref().map_or(false, |t| !t.is_empty());

    if !has_align && !has_line && !has_indent && !rtl && !has_tab_stops {
        // Minimal pPr: just no-bullet marker
        s.push_str("<a:pPr indent=\"0\" marL=\"0\"><a:buNone/></a:pPr>");
        return s;
    }

    s.push_str("<a:pPr");
    if rtl { s.push_str(" rtl=\"1\""); }
    if let Some(ref align) = opts.align {
        s.push_str(&format!(" algn=\"{}\"", align.as_ooxml()));
    }
    if let Some(level) = opts.indent_level {
        if level > 0 { s.push_str(&format!(" lvl=\"{level}\"")); }
    }
    s.push_str(" indent=\"0\" marL=\"0\">");

    if let Some(ls) = opts.line_spacing {
        s.push_str(&format!("<a:lnSpc><a:spcPts val=\"{}\"/></a:lnSpc>", (ls * 100.0).round() as i64));
    } else if let Some(lm) = opts.line_spacing_multiple {
        s.push_str(&format!("<a:lnSpc><a:spcPct val=\"{}\"/></a:lnSpc>", (lm * 100_000.0).round() as i64));
    }
    if let Some(before) = opts.para_space_before {
        s.push_str(&format!("<a:spcBef><a:spcPts val=\"{}\"/></a:spcBef>", (before * 100.0).round() as i64));
    }
    if let Some(after) = opts.para_space_after {
        s.push_str(&format!("<a:spcAft><a:spcPts val=\"{}\"/></a:spcAft>", (after * 100.0).round() as i64));
    }

    // Bullet
    match &opts.bullet {
        Some(bullet) => {
            match bullet.bullet_type {
                BulletType::Default => {
                    s.push_str("<a:buSzPct val=\"100000\"/><a:buChar char=\"&#x2022;\"/>");
                }
                BulletType::Numbered => {
                    let style = bullet.style.as_deref().unwrap_or("arabicPeriod");
                    let start = bullet.number_start_at.unwrap_or(1);
                    s.push_str(&format!("<a:buSzPct val=\"100000\"/><a:buFont typeface=\"+mj-lt\"/><a:buAutoNum type=\"{style}\" startAt=\"{start}\"/>"));
                }
                BulletType::Character => {
                    if let Some(ref code) = bullet.character_code {
                        s.push_str(&format!("<a:buSzPct val=\"100000\"/><a:buChar char=\"&#{code};\"/>"));
                    }
                }
            }
        }
        None => {
            s.push_str("<a:buNone/>");
        }
    }

    // Tab stops — must be wrapped in <a:tabLst>, after bullet elements
    if let Some(ref stops) = opts.tab_stops {
        if !stops.is_empty() {
            s.push_str("<a:tabLst>");
            for stop in stops {
                let pos = inch_to_emu(stop.pos_inches);
                let algn = &stop.align;
                s.push_str(&format!("<a:tab pos=\"{pos}\" algn=\"{algn}\"/>"));
            }
            s.push_str("</a:tabLst>");
        }
    }

    s.push_str("</a:pPr>");
    s
}

pub fn gen_xml_text_run(run: &TextRun, para_opts: &TextOptions) -> String {
    let run_opts = &run.options;
    let mut s = String::new();

    // Equation runs emit pre-rendered OMML directly
    if let Some(ref omml) = run.equation_omml {
        s.push_str(omml);
        return s;
    }

    // Field runs use <a:fld> instead of <a:r>
    if let Some(ref fld) = run.field {
        let fld_type = fld.as_ooxml();
        // Deterministic GUID based on field type string
        let guid = format!("{{B0E4D6F0-0000-0000-0000-{:0>12X}}}", {
            let mut hash: u64 = 5381;
            for b in fld_type.bytes() { hash = hash.wrapping_mul(33).wrapping_add(b as u64); }
            hash & 0xFFFF_FFFF_FFFF
        });
        s.push_str(&format!("<a:fld id=\"{guid}\" type=\"{fld_type}\">"));
        s.push_str("<a:rPr lang=\"en-US\" smtClean=\"0\"/>");
        let placeholder = if run.text.is_empty() { "\u{2039}#\u{203A}" } else { &run.text };
        s.push_str(&format!("<a:t>{}</a:t>", encode_xml_entities(placeholder)));
        s.push_str("</a:fld>");
        return s;
    }

    s.push_str("<a:r>");

    // Build <a:rPr> attributes
    let lang = run_opts.lang.as_deref().unwrap_or("en-US");
    let mut rpr_attrs = format!("<a:rPr lang=\"{lang}\"");

    let font_size = run_opts.font_size.or(para_opts.font_size);
    if let Some(fs) = font_size {
        rpr_attrs.push_str(&format!(" sz=\"{}\"", (fs * 100.0).round() as i64));
    }
    let bold = run_opts.bold.or(para_opts.bold);
    if let Some(b) = bold { rpr_attrs.push_str(&format!(" b=\"{}\"", if b { 1 } else { 0 })); }
    let italic = run_opts.italic.or(para_opts.italic);
    if let Some(i) = italic { rpr_attrs.push_str(&format!(" i=\"{}\"", if i { 1 } else { 0 })); }
    if let Some(ref u) = run_opts.underline { rpr_attrs.push_str(&format!(" u=\"{u}\"")); }
    if let Some(ref st) = run_opts.strike {
        let val = if st == "dbl" { "dblStrike" } else { "sngStrike" };
        rpr_attrs.push_str(&format!(" strike=\"{val}\""));
    }
    if run_opts.superscript { rpr_attrs.push_str(" baseline=\"30000\""); }
    else if run_opts.subscript { rpr_attrs.push_str(" baseline=\"-40000\""); }
    if let Some(cs) = run_opts.char_spacing {
        rpr_attrs.push_str(&format!(" spc=\"{}\" kern=\"0\"", (cs * 100.0).round() as i64));
    }
    rpr_attrs.push_str(" dirty=\"0\"");

    // Build <a:rPr> children — OOXML CT_TextCharacterProperties child order:
    // 1. <a:ln>   — text outline/stroke
    // 2. fill     — text color
    // 3. <a:effectLst> — effects (glow, shadow, etc.)
    // 4. <a:highlight> — highlight background
    // 5. <a:latin>/<a:ea>/<a:cs> — font faces
    // 6. <a:hlinkClick> — hyperlink
    let mut rpr_children = String::new();

    if let Some(ref outline) = run_opts.outline {
        let w = val_to_pts(outline.size);
        let c = outline.color.trim_start_matches('#').to_uppercase();
        rpr_children.push_str(&format!("<a:ln w=\"{w}\"><a:solidFill><a:srgbClr val=\"{c}\"/></a:solidFill></a:ln>"));
    }
    if let Some(ref uc) = run_opts.underline_color {
        let c = uc.trim_start_matches('#').to_uppercase();
        rpr_children.push_str(&format!("<a:uFill><a:solidFill><a:srgbClr val=\"{c}\"/></a:solidFill></a:uFill>"));
    }
    let color = run_opts.color.as_deref().or(para_opts.color.as_deref());
    if let Some(c) = color {
        rpr_children.push_str(&gen_xml_color_selection_str(c, run_opts.transparency));
    } else if let Some(t) = run_opts.transparency {
        // No explicit color but transparency requested — apply alpha to a transparent fill
        let alpha = ((100.0 - t) * 1000.0).round() as i64;
        rpr_children.push_str(&format!(
            "<a:solidFill><a:srgbClr val=\"000000\"><a:alpha val=\"{alpha}\"/></a:srgbClr></a:solidFill>"
        ));
    }
    if let Some(ref glow) = run_opts.glow {
        rpr_children.push_str(&format!("<a:effectLst>{}</a:effectLst>",
            create_glow_element(glow.size, &glow.color, glow.opacity)));
    }
    if let Some(ref hl_color) = run_opts.highlight {
        let c = hl_color.trim_start_matches('#').to_uppercase();
        rpr_children.push_str(&format!("<a:highlight><a:srgbClr val=\"{c}\"/></a:highlight>"));
    }
    let font_face = run_opts.font_face.as_deref().or(para_opts.font_face.as_deref());
    if let Some(ff) = font_face {
        rpr_children.push_str(&format!(
            "<a:latin typeface=\"{ff}\" pitchFamily=\"34\" charset=\"0\"/>\
<a:ea typeface=\"{ff}\" pitchFamily=\"34\" charset=\"-122\"/>\
<a:cs typeface=\"{ff}\" pitchFamily=\"34\" charset=\"-120\"/>"
        ));
    }
    if let Some(ref hl) = run_opts.hyperlink {
        let tooltip = hl.tooltip.as_deref().unwrap_or("");
        let tooltip_attr = if tooltip.is_empty() { String::new() } else { format!(" tooltip=\"{}\"", encode_xml_entities(tooltip)) };
        if let Some(ref nav) = hl.action {
            rpr_children.push_str(&format!("<a:hlinkClick r:id=\"\" action=\"{}\"{tooltip_attr}/>", nav.as_ppaction()));
        } else if hl.r_id > 0 {
            if hl.slide.is_some() {
                rpr_children.push_str(&format!(
                    "<a:hlinkClick r:id=\"rId{}\" action=\"ppaction://hlinksldjump\"{tooltip_attr}/>",
                    hl.r_id
                ));
            } else if hl.url.is_some() {
                rpr_children.push_str(&format!(
                    "<a:hlinkClick r:id=\"rId{}\"{tooltip_attr}/>",
                    hl.r_id
                ));
            }
        }
    }

    // Emit <a:rPr> — self-closing when no children
    if rpr_children.is_empty() {
        s.push_str(&rpr_attrs);
        s.push_str("/>");
    } else {
        s.push_str(&rpr_attrs);
        s.push('>');
        s.push_str(&rpr_children);
        s.push_str("</a:rPr>");
    }

    s.push_str(&format!("<a:t>{}</a:t>", encode_xml_entities(&run.text)));
    s.push_str("</a:r>");
    s
}

// ─────────────────────────────────────────────────────────────
// Slide number placeholder
// ─────────────────────────────────────────────────────────────

fn gen_xml_slide_number(sn: &crate::types::SlideNumberProps, obj_id: usize, pres: &Presentation) -> String {
    let layout = &pres.layout;
    let x = sn.x.as_ref().map(|c| c.to_emu(layout.width)).unwrap_or_else(|| inch_to_emu(0.1));
    let y = sn.y.as_ref().map(|c| c.to_emu(layout.height)).unwrap_or_else(|| inch_to_emu(7.0));
    let cx = sn.w.as_ref().map(|c| c.to_emu(layout.width)).unwrap_or_else(|| inch_to_emu(1.0));
    let cy = sn.h.as_ref().map(|c| c.to_emu(layout.height)).unwrap_or_else(|| inch_to_emu(0.4));

    // Build rPr attributes
    let mut rpr_attrs = String::from("<a:rPr lang=\"en-US\"");
    if let Some(fs) = sn.font_size {
        rpr_attrs.push_str(&format!(" sz=\"{}\"", (fs * 100.0).round() as i64));
    }
    if sn.bold { rpr_attrs.push_str(" b=\"1\""); }
    rpr_attrs.push_str(" smtClean=\"0\" dirty=\"0\"");

    // rPr children: color
    let mut rpr_children = String::new();
    if let Some(ref c) = sn.color {
        rpr_children.push_str(&gen_xml_color_selection_str(c, None));
    }
    if let Some(ref ff) = sn.font_face {
        rpr_children.push_str(&format!(
            "<a:latin typeface=\"{ff}\" pitchFamily=\"34\" charset=\"0\"/>"
        ));
    }

    let rpr = if rpr_children.is_empty() {
        format!("{rpr_attrs}/>")
    } else {
        format!("{rpr_attrs}>{rpr_children}</a:rPr>")
    };

    let align_attr = sn.align.as_deref().map(|a| format!(" algn=\"{a}\"")).unwrap_or_default();

    format!(
        "<p:sp>\
<p:nvSpPr>\
<p:cNvPr id=\"{obj_id}\" name=\"Slide Number Placeholder {obj_id}\"/>\
<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>\
<p:nvPr><p:ph type=\"sldNum\" sz=\"quarter\" idx=\"12\"/></p:nvPr>\
</p:nvSpPr>\
<p:spPr>\
<a:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>\
<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>\
</p:spPr>\
<p:txBody>\
<a:bodyPr/>\
<a:lstStyle/>\
<a:p>\
<a:pPr{align_attr}/>\
<a:fld id=\"{{F7021451-1387-4CA6-816F-3879F97B5CBC}}\" type=\"slidenum\">\
{rpr}\
<a:t>‹#›</a:t>\
</a:fld>\
</a:p>\
</p:txBody>\
</p:sp>"
    )
}

// ─────────────────────────────────────────────────────────────
// Custom geometry XML helper
// ─────────────────────────────────────────────────────────────

fn gen_xml_custom_geom(pts: &[crate::types::CustomGeomPoint], _cx: i64, _cy: i64) -> String {
    use crate::types::CustomGeomPoint as CGP;
    // Use a 100000×100000 normalised coordinate space; fractions map linearly.
    const W: i64 = 100_000;
    const H: i64 = 100_000;
    let px = |x: f64| (x * W as f64).round() as i64;
    let py = |y: f64| (y * H as f64).round() as i64;

    let mut ops = String::new();
    for pt in pts {
        match pt {
            CGP::MoveTo(x, y) => {
                ops.push_str(&format!("<a:moveTo><a:pt x=\"{}\" y=\"{}\"/></a:moveTo>", px(*x), py(*y)));
            }
            CGP::LineTo(x, y) => {
                ops.push_str(&format!("<a:lnTo><a:pt x=\"{}\" y=\"{}\"/></a:lnTo>", px(*x), py(*y)));
            }
            CGP::ArcTo { w_r, h_r, start_angle, swing_angle } => {
                let wr = px(*w_r);
                let hr = py(*h_r);
                let st = (*start_angle * 60_000.0).round() as i64;
                let sw = (*swing_angle * 60_000.0).round() as i64;
                ops.push_str(&format!("<a:arcTo wR=\"{wr}\" hR=\"{hr}\" stAng=\"{st}\" swAng=\"{sw}\"/>"));
            }
            CGP::CubicBezTo(cp1x, cp1y, cp2x, cp2y, ex, ey) => {
                ops.push_str(&format!(
                    "<a:cubicBezTo>\
<a:pt x=\"{}\" y=\"{}\"/>\
<a:pt x=\"{}\" y=\"{}\"/>\
<a:pt x=\"{}\" y=\"{}\"/>\
</a:cubicBezTo>",
                    px(*cp1x), py(*cp1y), px(*cp2x), py(*cp2y), px(*ex), py(*ey)
                ));
            }
            CGP::QuadBezTo(cpx, cpy, ex, ey) => {
                ops.push_str(&format!(
                    "<a:quadBezTo>\
<a:pt x=\"{}\" y=\"{}\"/>\
<a:pt x=\"{}\" y=\"{}\"/>\
</a:quadBezTo>",
                    px(*cpx), py(*cpy), px(*ex), py(*ey)
                ));
            }
            CGP::Close => {
                ops.push_str("<a:close/>");
            }
        }
    }

    format!(
        "<a:custGeom>\
<a:avLst/>\
<a:gdLst/>\
<a:ahLst/>\
<a:cxnLst/>\
<a:rect l=\"0\" t=\"0\" r=\"{W}\" b=\"{H}\"/>\
<a:pathLst>\
<a:path w=\"{W}\" h=\"{H}\">{ops}</a:path>\
</a:pathLst>\
</a:custGeom>"
    )
}

// ─────────────────────────────────────────────────────────────
// Shape line XML helper (reused by shapes and text boxes)
// ─────────────────────────────────────────────────────────────

fn gen_xml_shape_line(line: &crate::types::ShapeLineProps) -> String {
    let width_attr = line.width
        .map(|w| format!(" w=\"{}\"", val_to_pts(w)))
        .unwrap_or_default();
    let cap_attr = line.cap
        .map(|c| format!(" cap=\"{}\"", c.as_str()))
        .unwrap_or_default();
    let mut s = format!("<a:ln{width_attr}{cap_attr}>");
    if let Some(ref color) = line.color {
        s.push_str(&gen_xml_color_selection_str(color, line.transparency));
    }
    if let Some(ref dash) = line.dash_type {
        s.push_str(&format!("<a:prstDash val=\"{dash}\"/>"));
    }
    // Line join style
    match line.join {
        Some(crate::types::LineJoin::Round) => s.push_str("<a:round/>"),
        Some(crate::types::LineJoin::Bevel) => s.push_str("<a:bevel/>"),
        Some(crate::types::LineJoin::Miter) => s.push_str("<a:miter/>"),
        None => {}
    }
    if let Some(ref arrow) = line.begin_arrow_type {
        s.push_str(&format!("<a:headEnd type=\"{arrow}\"/>"));
    }
    if let Some(ref arrow) = line.end_arrow_type {
        s.push_str(&format!("<a:tailEnd type=\"{arrow}\"/>"));
    }
    s.push_str("</a:ln>");
    s
}

// ─────────────────────────────────────────────────────────────
// Connector object → <p:cxnSp>
// ─────────────────────────────────────────────────────────────

fn gen_xml_connector_object(c: &ConnectorObject, obj_id: usize, pres: &Presentation) -> String {
    let layout = &pres.layout;

    // Resolve endpoint coordinates to EMU
    let x1 = c.options.x1.as_ref().map(|v| v.to_emu(layout.width)).unwrap_or(0);
    let y1 = c.options.y1.as_ref().map(|v| v.to_emu(layout.height)).unwrap_or(0);
    let x2 = c.options.x2.as_ref().map(|v| v.to_emu(layout.width)).unwrap_or(0);
    let y2 = c.options.y2.as_ref().map(|v| v.to_emu(layout.height)).unwrap_or(0);

    // Bounding box: top-left corner + dimensions
    let bx = x1.min(x2);
    let by = y1.min(y2);
    let cx = (x2 - x1).unsigned_abs() as i64;
    let cy = (y2 - y1).unsigned_abs() as i64;

    // Flip flags when end is to the left or above start
    let flip_h = if x2 < x1 { " flipH=\"1\"" } else { "" };
    let flip_v = if y2 < y1 { " flipV=\"1\"" } else { "" };

    let prst = c.connector_type.as_prst();
    let obj_name = &c.object_name;

    // Optional connection site elements
    let st_cxn = if let Some(ref ep) = c.options.start_conn {
        format!("<a:stCxn id=\"{}\" idx=\"{}\"/>", ep.shape_id, ep.site_idx)
    } else {
        String::new()
    };
    let end_cxn = if let Some(ref ep) = c.options.end_conn {
        format!("<a:endCxn id=\"{}\" idx=\"{}\"/>", ep.shape_id, ep.site_idx)
    } else {
        String::new()
    };

    let mut s = String::new();
    s.push_str(&format!(
        "<p:cxnSp>\
<p:nvCxnSpPr>\
<p:cNvPr id=\"{obj_id}\" name=\"{obj_name}\"/>\
<p:cNvCxnSpPr>{st_cxn}{end_cxn}</p:cNvCxnSpPr>\
<p:nvPr/>\
</p:nvCxnSpPr>\
<p:spPr>\
<a:xfrm{flip_h}{flip_v}>\
<a:off x=\"{bx}\" y=\"{by}\"/>\
<a:ext cx=\"{cx}\" cy=\"{cy}\"/>\
</a:xfrm>\
<a:prstGeom prst=\"{prst}\"><a:avLst/></a:prstGeom>\
<a:noFill/>"
    ));

    // Line styling
    if let Some(ref line) = c.options.line {
        let width_attr = line.width
            .map(|w| format!(" w=\"{}\"", val_to_pts(w)))
            .unwrap_or_default();
        s.push_str(&format!("<a:ln{width_attr}>"));
        if let Some(ref color) = line.color {
            s.push_str(&gen_xml_color_selection_str(color, line.transparency));
        }
        if let Some(ref dash) = line.dash_type {
            s.push_str(&format!("<a:prstDash val=\"{dash}\"/>"));
        }
        if let Some(ref arrow) = line.begin_arrow_type {
            s.push_str(&format!("<a:headEnd type=\"{arrow}\"/>"));
        }
        if let Some(ref arrow) = line.end_arrow_type {
            s.push_str(&format!("<a:tailEnd type=\"{arrow}\"/>"));
        }
        s.push_str("</a:ln>");
    }

    s.push_str("</p:spPr>");
    s.push_str("</p:cxnSp>");
    s
}

// ─────────────────────────────────────────────────────────────
// Slide transition XML
// ─────────────────────────────────────────────────────────────

pub fn gen_xml_transition(slide: &Slide) -> String {
    use crate::types::{TransitionType};

    let props = match &slide.transition {
        Some(p) => p,
        None => return String::new(),
    };

    if props.transition_type == TransitionType::None {
        return String::new();
    }

    let mut attrs = String::new();

    // Speed / duration
    if let Some(ms) = props.duration_ms {
        attrs.push_str(&format!(" dur=\"{ms}\""));
    } else if let Some(ref spd) = props.speed {
        attrs.push_str(&format!(" spd=\"{}\"", spd.as_str()));
    }

    // Auto-advance
    if let Some(ms) = props.advance_after_ms {
        attrs.push_str(&format!(" advTm=\"{ms}\""));
    }

    // advClick="1" is the OOXML default — only emit when explicitly disabled
    if !props.advance_on_click {
        attrs.push_str(" advClick=\"0\"");
    }

    // Direction attribute helper — skip "l" (left) since it is the default for all directional transitions
    let dir_attr = props.direction.as_ref()
        .filter(|d| !matches!(d, TransitionDir::Left))
        .map(|d| format!(" dir=\"{}\"", d.as_str()))
        .unwrap_or_default();

    // Inner transition element
    let inner = match &props.transition_type {
        TransitionType::None => return String::new(),
        TransitionType::Cut    => "<p:cut/>".to_string(),
        TransitionType::Fade   => "<p:fade/>".to_string(),
        TransitionType::Flash  => "<p:flash/>".to_string(),
        TransitionType::Morph  => "<p:morph/>".to_string(),
        TransitionType::Zoom   => "<p:zoom/>".to_string(),
        TransitionType::Vortex => "<p:vortex/>".to_string(),
        TransitionType::Ripple => "<p:ripple/>".to_string(),
        TransitionType::Glitter=> "<p:glitter/>".to_string(),
        TransitionType::Honeycomb => "<p:honeycomb/>".to_string(),
        TransitionType::Shred  => "<p:shred/>".to_string(),
        TransitionType::Switch => "<p:switch/>".to_string(),
        TransitionType::Flip   => "<p:flip/>".to_string(),
        TransitionType::Pan    => format!("<p:pan{dir_attr}/>"),
        TransitionType::Ferris => format!("<p:ferris{dir_attr}/>"),
        TransitionType::Gallery=> format!("<p:gallery{dir_attr}/>"),
        TransitionType::Conveyor=> format!("<p:conveyor{dir_attr}/>"),
        TransitionType::Doors  => format!("<p:doors{dir_attr}/>"),
        TransitionType::Box    => format!("<p:box{dir_attr}/>"),
        TransitionType::Push   => format!("<p:push{dir_attr}/>"),
        TransitionType::Wipe   => format!("<p:wipe{dir_attr}/>"),
        TransitionType::Cover  => format!("<p:cover{dir_attr}/>"),
        TransitionType::Uncover=> format!("<p:uncover{dir_attr}/>"),
        TransitionType::Random => "<p:random/>".to_string(),
        TransitionType::RandomBar => {
            let vert = matches!(props.direction, Some(TransitionDir::Up) | Some(TransitionDir::Down));
            format!("<p:randomBar dir=\"{}\"/>", if vert { "vert" } else { "horz" })
        }
        TransitionType::Circle  => "<p:circle/>".to_string(),
        TransitionType::Diamond => "<p:diamond/>".to_string(),
        TransitionType::Wheel   => "<p:wheel spokes=\"4\"/>".to_string(),
        TransitionType::Checker => {
            let vert = matches!(props.direction, Some(TransitionDir::Up) | Some(TransitionDir::Down));
            format!("<p:checkerboard dir=\"{}\"/>", if vert { "vert" } else { "horz" })
        }
        TransitionType::Blinds  => {
            let vert = matches!(props.direction, Some(TransitionDir::Up) | Some(TransitionDir::Down));
            format!("<p:blinds dir=\"{}\"/>", if vert { "vert" } else { "horz" })
        }
        TransitionType::Strips  => format!("<p:strips{dir_attr}/>"),
        TransitionType::Plus    => "<p:plus/>".to_string(),
        TransitionType::Split   => {
            // Split needs orientation: horz/vert and in/out
            format!("<p:split orient=\"horz\" dir=\"in\"/>")
        }
    };

    format!("<p:transition{attrs}>{inner}</p:transition>")
}

// ─────────────────────────────────────────────────────────────
// Pattern fill XML helper
// ─────────────────────────────────────────────────────────────

fn gen_xml_pattern_fill(pf: &PatternFill) -> String {
    let prst = pf.pattern.as_str();
    let fg = pf.fg_color.trim_start_matches('#').to_uppercase();
    let bg = pf.bg_color.trim_start_matches('#').to_uppercase();
    format!(
        "<a:pattFill prst=\"{prst}\">\
<a:fgClr><a:srgbClr val=\"{fg}\"/></a:fgClr>\
<a:bgClr><a:srgbClr val=\"{bg}\"/></a:bgClr>\
</a:pattFill>"
    )
}

// ─────────────────────────────────────────────────────────────
// Shadow XML helper
// ─────────────────────────────────────────────────────────────

fn gen_xml_shadow(shadow: &crate::types::ShadowProps) -> String {
    let blur = val_to_pts(shadow.blur.unwrap_or(8.0));
    let offset = val_to_pts(shadow.offset.unwrap_or(4.0));
    let angle = (shadow.angle.unwrap_or(270.0) * 60_000.0).round() as i64;
    let opacity = ((shadow.opacity.unwrap_or(0.75)) * 100_000.0).round() as i64;
    let color = shadow.color.as_deref().unwrap_or("000000");

    let type_tag = match shadow.shadow_type {
        ShadowType::Outer => "outerShdw",
        ShadowType::Inner => "innerShdw",
        ShadowType::None => return String::new(),
    };
    // For outerShdw: omit sx/sy/kx/ky (defaults) and rotWithShape (PowerPoint strips it on repair)
    let outer_attrs = if shadow.shadow_type == ShadowType::Outer {
        String::from(" algn=\"bl\"")
    } else {
        String::new()
    };

    format!(
        "<a:effectLst><a:{type_tag}{outer_attrs} blurRad=\"{blur}\" dist=\"{offset}\" dir=\"{angle}\">\
<a:srgbClr val=\"{color}\"><a:alpha val=\"{opacity}\"/></a:srgbClr>\
</a:{type_tag}></a:effectLst>",
        type_tag = type_tag,
        outer_attrs = outer_attrs,
        blur = blur,
        offset = offset,
        angle = angle,
        color = color,
        opacity = opacity,
    )
}

// ─────────────────────────────────────────────────────────────
// 3D effect helpers
// ─────────────────────────────────────────────────────────────

fn gen_xml_scene3d(scene: &crate::types::Scene3DProps) -> String {
    let cam = &scene.camera;
    let lr = &scene.light_rig;
    let mut s = String::from("<a:scene3d>");

    // Camera
    let fov_attr = cam.fov.map(|f| format!(" fov=\"{f}\"")).unwrap_or_default();
    s.push_str(&format!("<a:camera prst=\"{}\"{fov_attr}", cam.preset.as_str()));
    if let Some(ref rot) = cam.rotation {
        s.push_str(&format!("><a:rot lat=\"{}\" lon=\"{}\" rev=\"{}\"/></a:camera>", rot.lat, rot.lon, rot.rev));
    } else {
        s.push_str("/>");
    }

    // Light rig
    s.push_str(&format!("<a:lightRig rig=\"{}\" dir=\"{}\"", lr.rig_type.as_str(), lr.direction.as_str()));
    if let Some(ref rot) = lr.rotation {
        s.push_str(&format!("><a:rot lat=\"{}\" lon=\"{}\" rev=\"{}\"/></a:lightRig>", rot.lat, rot.lon, rot.rev));
    } else {
        s.push_str("/>");
    }

    s.push_str("</a:scene3d>");
    s
}

fn gen_xml_sp3d(sp3d: &crate::types::Shape3DProps) -> String {
    let mut attrs = String::new();
    if let Some(h) = sp3d.extrusion_height {
        attrs.push_str(&format!(" extrusionH=\"{h}\""));
    }
    if let Some(w) = sp3d.contour_width {
        attrs.push_str(&format!(" contourW=\"{w}\""));
    }
    if let Some(ref mat) = sp3d.material {
        attrs.push_str(&format!(" prstMaterial=\"{}\"", mat.as_str()));
    }
    let mut s = format!("<a:sp3d{attrs}>");

    if let Some(ref bevel) = sp3d.bevel_top {
        let w = bevel.width.unwrap_or(76200);
        let h = bevel.height.unwrap_or(76200);
        s.push_str(&format!("<a:bevelT w=\"{w}\" h=\"{h}\" prst=\"{}\"/>", bevel.preset.as_str()));
    }
    if let Some(ref bevel) = sp3d.bevel_bottom {
        let w = bevel.width.unwrap_or(76200);
        let h = bevel.height.unwrap_or(76200);
        s.push_str(&format!("<a:bevelB w=\"{w}\" h=\"{h}\" prst=\"{}\"/>", bevel.preset.as_str()));
    }
    if let Some(ref color) = sp3d.contour_color {
        s.push_str(&format!("<a:contourClr><a:srgbClr val=\"{color}\"/></a:contourClr>"));
    }

    s.push_str("</a:sp3d>");
    s
}

// ─────────────────────────────────────────────────────────────
// Position / transform helpers
// ─────────────────────────────────────────────────────────────

fn resolve_position(pos: &crate::types::PositionProps, pres: &Presentation) -> (i64, i64, i64, i64) {
    let layout = &pres.layout;
    let x = pos.x.as_ref().map(|c| c.to_emu(layout.width)).unwrap_or(EMU);
    let y = pos.y.as_ref().map(|c| c.to_emu(layout.height)).unwrap_or(EMU);
    let cx = pos.w.as_ref().map(|c| c.to_emu(layout.width))
        .unwrap_or_else(|| (layout.width as f64 * 0.75).round() as i64);
    let cy = pos.h.as_ref().map(|c| c.to_emu(layout.height)).unwrap_or(EMU);
    (x, y, cx, cy)
}

fn build_location_attr(rotate: Option<f64>, flip_h: bool, flip_v: bool) -> String {
    let mut s = String::new();
    if let Some(r) = rotate {
        s.push_str(&format!(" rot=\"{}\"", convert_rotation_degrees(r)));
    }
    if flip_h { s.push_str(" flipH=\"1\""); }
    if flip_v { s.push_str(" flipV=\"1\""); }
    s
}

// ─────────────────────────────────────────────────────────────
// Animation timing block
// ─────────────────────────────────────────────────────────────

/// Generate `<p:timing>` for all animated objects on the slide.
/// Returns an empty string when no objects have animations.
pub fn gen_xml_timing(objects: &[SlideObject]) -> String {
    use crate::types::AnimationEffect;

    // Collect (spid, &AnimationEffect) — spid = cNvPr id = idx + 2.
    // Each object can carry multiple animations; flatten them, preserving spid.
    let anims: Vec<(usize, &AnimationEffect)> = objects
        .iter()
        .enumerate()
        .flat_map(|(idx, obj)| {
            let empty: &[AnimationEffect] = &[];
            let anim_slice: &[AnimationEffect] = match obj {
                SlideObject::Text(t)      => &t.options.animations,
                SlideObject::Shape(s)     => &s.options.animations,
                SlideObject::Image(i)     => &i.options.animations,
                SlideObject::Table(t)     => &t.options.animations,
                SlideObject::Connector(c) => &c.options.animations,
                SlideObject::Media(_)     => empty,
                SlideObject::Group(_)     => empty,
            };
            let spid = idx + 2;
            anim_slice.iter().map(move |a| (spid, a))
        })
        .collect();

    if anims.is_empty() { return String::new(); }

    use crate::types::AnimationTrigger;

    // ── Group animations by trigger type ─────────────────────────────────────
    // Each "trigger group" becomes one outer <p:par> in the mainSeq.
    //  - OnClick (no click_group): singleton group, outer delay="indefinite"
    //  - OnClick (with click_group): merged by key, outer delay="indefinite"
    //  - WithPrevious: always a new group, outer delay="0" (or delay_ms)
    //  - AfterPrevious: always a new group, outer delay="0" (or delay_ms)
    #[derive(Clone, Copy, PartialEq)]
    enum TriggerKind { Click, With, After }

    struct AnimGroup<'a> {
        trigger: TriggerKind,
        delay_ms: u32,
        items: Vec<(usize, &'a AnimationEffect)>,
    }

    let mut groups: Vec<AnimGroup<'_>> = Vec::new();
    let mut click_key_to_idx: Vec<(u32, usize)> = Vec::new();

    for (spid, anim) in &anims {
        match &anim.trigger {
            AnimationTrigger::OnClick => {
                match anim.click_group {
                    None => groups.push(AnimGroup {
                        trigger: TriggerKind::Click,
                        delay_ms: anim.delay_ms.unwrap_or(0),
                        items: vec![(*spid, *anim)],
                    }),
                    Some(key) => {
                        if let Some(&(_, gi)) = click_key_to_idx.iter().find(|(k, _)| *k == key) {
                            groups[gi].items.push((*spid, *anim));
                        } else {
                            let gi = groups.len();
                            click_key_to_idx.push((key, gi));
                            groups.push(AnimGroup {
                                trigger: TriggerKind::Click,
                                delay_ms: 0,
                                items: vec![(*spid, *anim)],
                            });
                        }
                    }
                }
            }
            AnimationTrigger::WithPrevious => {
                groups.push(AnimGroup {
                    trigger: TriggerKind::With,
                    delay_ms: anim.delay_ms.unwrap_or(0),
                    items: vec![(*spid, *anim)],
                });
            }
            AnimationTrigger::AfterPrevious => {
                groups.push(AnimGroup {
                    trigger: TriggerKind::After,
                    delay_ms: anim.delay_ms.unwrap_or(0),
                    items: vec![(*spid, *anim)],
                });
            }
        }
    }

    let mut s = String::new();
    s.push_str("<p:timing><p:tnLst><p:par>");
    s.push_str("<p:cTn id=\"1\" dur=\"indefinite\" restart=\"never\" nodeType=\"tmRoot\"><p:childTnLst>");
    s.push_str("<p:seq concurrent=\"1\" nextAc=\"seek\">");
    s.push_str("<p:cTn id=\"2\" dur=\"indefinite\" nodeType=\"mainSeq\"><p:childTnLst>");

    let mut bld_entries: Vec<(usize, usize, &'static str, bool)> = Vec::new(); // (spid, grpId, presetClass, has_sub_target)
    let mut ctn_id: usize = 3; // running unique cTn id counter (1=tmRoot, 2=mainSeq reserved)

    for (grp_idx, group) in groups.iter().enumerate() {
        let id_outer  = ctn_id; ctn_id += 1;
        let id_middle = ctn_id; ctn_id += 1;

        // Outer par delay: OnClick waits for a click; With/After start automatically.
        let outer_delay = match group.trigger {
            TriggerKind::Click => "indefinite".to_string(),
            _ => group.delay_ms.to_string(),
        };

        // For OnClick with delay_ms, the delay goes on the inner (middle) par.
        let middle_delay = match group.trigger {
            TriggerKind::Click => group.delay_ms.to_string(),
            _ => "0".to_string(),
        };

        s.push_str(&format!(
            "<p:par><p:cTn id=\"{id_outer}\" fill=\"hold\">\
<p:stCondLst><p:cond delay=\"{outer_delay}\"/></p:stCondLst>\
<p:childTnLst><p:par><p:cTn id=\"{id_middle}\" fill=\"hold\">\
<p:stCondLst><p:cond delay=\"{middle_delay}\"/></p:stCondLst>\
<p:childTnLst>"
        ));

        for (j, (spid, anim)) in group.items.iter().enumerate() {
            let id_effect = ctn_id; ctn_id += 1;
            let id_set    = ctn_id; ctn_id += 1;
            let id_fx     = ctn_id; ctn_id += 1;
            let id_fx2    = ctn_id; ctn_id += 1;

            let (preset_id, preset_class, preset_subtype, mut inner_xml) =
                build_effect_xml(&anim.effect, *spid, id_set, id_fx, id_fx2);

            // Rewrite target for text sub-range if specified
            if let Some(tt) = &anim.text_target {
                let bare = format!("<p:spTgt spid=\"{spid}\"/>");
                let with_tx = match tt {
                    crate::types::TextTarget::CharRange { st_idx, end_idx } =>
                        format!("<p:spTgt spid=\"{spid}\"><p:txEl>\
                                 <p:charRg st=\"{st_idx}\" end=\"{end_idx}\"/>\
                                 </p:txEl></p:spTgt>"),
                    crate::types::TextTarget::ParaRange { st_idx, end_idx } =>
                        format!("<p:spTgt spid=\"{spid}\"><p:txEl>\
                                 <p:pRg st=\"{st_idx}\" end=\"{end_idx}\"/>\
                                 </p:txEl></p:spTgt>"),
                };
                inner_xml = inner_xml.replace(&bare, &with_tx);
            }

            // Determine nodeType based on trigger and position within group
            let node_type = match group.trigger {
                TriggerKind::Click => if j == 0 { "clickEffect" } else { "withEffect" },
                TriggerKind::With  => "withEffect",
                TriggerKind::After => "afterEffect",
            };

            // animBg="1" applies when the whole paragraph/shape is built; omit it for
            // character-range / paragraph-range sub-targets (charRg / pRg) because the
            // OOXML spec states animBg must NOT be set when animating selected characters.
            let has_sub_target = anim.text_target.is_some();
            bld_entries.push((*spid, grp_idx, preset_class, has_sub_target));

            s.push_str(&format!(
                "<p:par>\
<p:cTn id=\"{id_effect}\" presetID=\"{preset_id}\" presetClass=\"{preset_class}\" \
presetSubtype=\"{preset_subtype}\" fill=\"hold\" grpId=\"{grp_idx}\" nodeType=\"{node_type}\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst>\
<p:childTnLst>{inner_xml}</p:childTnLst></p:cTn>\
</p:par>"
            ));
        }

        // Close outer par
        s.push_str("</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>");
    }

    s.push_str("</p:childTnLst></p:cTn>");
    s.push_str("<p:prevCondLst><p:cond evt=\"onPrev\" delay=\"0\"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>");
    s.push_str("<p:nextCondLst><p:cond evt=\"onNext\" delay=\"0\"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>");
    s.push_str("</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>");

    s.push_str("<p:bldLst>");
    for (spid, grp_id, _preset_class, has_sub_target) in &bld_entries {
        // animBg="1" applies only for whole-shape/whole-paragraph animations.
        // Sub-target animations (charRg / pRg) must NOT have animBg set.
        if *has_sub_target {
            s.push_str(&format!("<p:bldP spid=\"{spid}\" grpId=\"{grp_id}\" uiExpand=\"1\"/>"));
        } else {
            s.push_str(&format!("<p:bldP spid=\"{spid}\" grpId=\"{grp_id}\" animBg=\"1\"/>"));
        }
    }
    s.push_str("</p:bldLst></p:timing>");
    s
}

/// Returns `(presetID, presetClass, presetSubtype, inner_xml)`.
///
/// `id_set` is used for the `<p:set>` visibility node (entrance/exit).
/// `id_fx`  is used for the motion/filter animation node.
///
/// Entrance effects:  `<p:set visibility="visible">` + animation  (shape is hidden at slide
///                    start because PowerPoint infers initial-hidden state from the set node)
/// Exit effects:      animation + `<p:set visibility="hidden" delay={dur}>` (hides permanently)
/// Emphasis effects:  animation only (shape stays visible throughout)
fn build_effect_xml(
    effect: &AnimationEffectType,
    spid: usize,
    id_set: usize,
    id_fx: usize,
    id_fx2: usize,
) -> (u32, &'static str, u32, String) {
    /// Prepend a make-visible set to an entrance animation XML string.
    fn entr(vis_xml: String, fx_xml: String) -> String { format!("{vis_xml}{fx_xml}") }
    /// Append a delayed make-hidden set to an exit animation XML string.
    fn exit_xml(fx_xml: String, hide_xml: String) -> String { format!("{fx_xml}{hide_xml}") }

    match effect {
        // ── Instant visibility toggle — no separate animation node ──────
        AnimationEffectType::Appear    => (1, "entr", 0, set_visibility_xml(spid, id_set, "visible", 0)),
        AnimationEffectType::Disappear => (1, "exit", 0, set_visibility_xml(spid, id_set, "hidden",  0)),

        // ── Fade ────────────────────────────────────────────────────────
        AnimationEffectType::FadeIn  => (10, "entr", 0,
            entr(set_visibility_xml(spid, id_set, "visible", 0),
                 anim_effect_filter(spid, id_fx, "in",  "fade", 500))),
        AnimationEffectType::FadeOut => (10, "exit", 0,
            exit_xml(anim_effect_filter(spid, id_fx, "out", "fade", 500),
                     set_visibility_xml(spid, id_set, "hidden", 500))),

        // ── Wipe ────────────────────────────────────────────────────────
        AnimationEffectType::WipeIn(dir) => {
            let (subtype, filter) = wipe_filter(dir);
            (4, "entr", subtype,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", &filter, 500)))
        }
        AnimationEffectType::WipeOut(dir) => {
            let (subtype, filter) = wipe_filter(dir);
            (4, "exit", subtype,
             exit_xml(anim_effect_filter(spid, id_fx, "out", &filter, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }

        // ── Split ───────────────────────────────────────────────────────
        AnimationEffectType::SplitIn(orient) => {
            let (subtype, filter) = split_filter(orient, true);
            (3, "entr", subtype,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", &filter, 500)))
        }
        AnimationEffectType::SplitOut(orient) => {
            let (subtype, filter) = split_filter(orient, false);
            (3, "exit", subtype,
             exit_xml(anim_effect_filter(spid, id_fx, "out", &filter, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }

        // ── Zoom ────────────────────────────────────────────────────────
        AnimationEffectType::ZoomIn  => (11, "entr", 0,
            entr(set_visibility_xml(spid, id_set, "visible", 0),
                 anim_scale_xml(spid, id_fx, 10000,  100000, 500))),
        AnimationEffectType::ZoomOut => (11, "exit", 0,
            exit_xml(anim_scale_xml(spid, id_fx, 100000, 10000, 500),
                     set_visibility_xml(spid, id_set, "hidden", 500))),

        // ── Fly ─────────────────────────────────────────────────────────
        AnimationEffectType::FlyIn(dir) => {
            let (subtype, attr, start_val, end_val) = fly_params(dir, true);
            (2, "entr", subtype,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_fly_xml(spid, id_fx, attr, start_val, end_val, 500)))
        }
        AnimationEffectType::FlyOut(dir) => {
            let (subtype, attr, start_val, end_val) = fly_params(dir, false);
            (2, "exit", subtype,
             exit_xml(anim_fly_xml(spid, id_fx, attr, start_val, end_val, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }

        // ── Blinds ──────────────────────────────────────────────────────
        AnimationEffectType::BlindsIn(orient) => {
            let (subtype, filter) = blinds_filter(orient);
            (3, "entr", subtype,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", &filter, 500)))
        }

        // ── Checkerboard ────────────────────────────────────────────────
        AnimationEffectType::CheckerboardIn(dir) => {
            let (subtype, filter) = checkerboard_filter(dir);
            (5, "entr", subtype,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", &filter, 500)))
        }

        // ── Dissolve In ─────────────────────────────────────────────────
        AnimationEffectType::DissolveIn =>
            (12, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", "dissolve()", 500))),

        // ── Peek In ─────────────────────────────────────────────────────
        AnimationEffectType::PeekIn(dir) => {
            let (subtype, filter) = wipe_filter(dir);
            (13, "entr", subtype,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", &filter, 500)))
        }

        // ── Random Bars ─────────────────────────────────────────────────
        AnimationEffectType::RandomBarsIn(orient) => {
            let (subtype, filter) = random_bars_filter(orient);
            (14, "entr", subtype,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", &filter, 500)))
        }

        // ── Shape (Box / Circle / Diamond / Plus) ───────────────────────
        AnimationEffectType::ShapeIn(variant) => {
            let (preset_id, filter) = shape_filter(variant, true);
            (preset_id, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", &filter, 500)))
        }

        // ── Strips ──────────────────────────────────────────────────────
        AnimationEffectType::StripsIn(dir) => {
            let (subtype, filter) = strips_filter(dir);
            (6, "entr", subtype,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", &filter, 500)))
        }

        // ── Wedge ───────────────────────────────────────────────────────
        AnimationEffectType::WedgeIn =>
            (17, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", "wedge()", 500))),

        // ── Wheel ───────────────────────────────────────────────────────
        AnimationEffectType::WheelIn(spokes) => {
            let n = (*spokes).max(1);
            let filter = format!("wheel(spokes={n})");
            (18, "entr", n,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_effect_filter(spid, id_fx, "in", &filter, 500)))
        }

        // ── Expand ──────────────────────────────────────────────────────
        AnimationEffectType::ExpandIn =>
            (22, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_scale_xml(spid, id_fx, 0, 100000, 500))),

        // ── Swivel ──────────────────────────────────────────────────────
        // Approximated as horizontal-axis swivel: width grows from 0 to 100%
        AnimationEffectType::SwivelIn =>
            (21, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_scale_xy_xml(spid, id_fx, 0, 100000, 100000, 100000, 500))),

        // ── Basic Zoom ──────────────────────────────────────────────────
        AnimationEffectType::BasicZoomIn =>
            (27, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_scale_xml(spid, id_fx, 10000, 100000, 500))),

        // ── Centre Revolve ──────────────────────────────────────────────
        // Grows from small while revolving — scale + rotation
        AnimationEffectType::CentreRevolveIn => {
            let rot_by = (720.0_f32 * 60_000.0).round() as i64; // 2 full rotations
            (23, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  format!("{}{}", anim_scale_xml(spid, id_fx, 10000, 100000, 700),
                          anim_rot_xml(spid, id_fx2, rot_by, 700))))
        }

        // ── Float In ────────────────────────────────────────────────────
        // Fades in while rotating from -90° and moving from upper-right (matches PowerPoint)
        AnimationEffectType::FloatIn(_dir) => {
            (30, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  format!("{}{}", anim_effect_filter(spid, id_fx, "in", "fade", 800),
                          anim_style_rotation_xml(spid, id_fx2, -90.0, 0.0, 800))))
        }

        // ── Grow Turn ───────────────────────────────────────────────────
        // Scales up while rotating 90°
        AnimationEffectType::GrowTurnIn => {
            let rot_by = (90.0_f32 * 60_000.0).round() as i64;
            (24, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  format!("{}{}", anim_scale_xml(spid, id_fx, 10000, 100000, 500),
                          anim_rot_xml(spid, id_fx2, rot_by, 500))))
        }

        // ── Rise Up ─────────────────────────────────────────────────────
        // Rises from below (position animation only)
        AnimationEffectType::RiseUpIn =>
            (25, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_fly_xml(spid, id_fx, "ppt_y", 0.5, 0.0, 500))),

        // ── Spinner ─────────────────────────────────────────────────────
        // Scales up while spinning a full rotation
        AnimationEffectType::SpinnerIn => {
            let rot_by = (360.0_f32 * 60_000.0).round() as i64;
            (28, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  format!("{}{}", anim_scale_xml(spid, id_fx, 10000, 100000, 700),
                          anim_rot_xml(spid, id_fx2, rot_by, 700))))
        }

        // ── Stretch ─────────────────────────────────────────────────────
        // Stretches in along one axis
        AnimationEffectType::StretchIn(dir) => {
            let (from_x, from_y, subtype) = match dir {
                Direction::Left | Direction::Right => (0, 100000, 4u32),
                Direction::Up   | Direction::Down  => (100000, 0, 8u32),
            };
            (29, "entr", subtype,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_scale_xy_xml(spid, id_fx, from_x, from_y, 100000, 100000, 500)))
        }

        // ── Boomerang ───────────────────────────────────────────────────
        // Approximated as fly-in from right with scale
        AnimationEffectType::BoomerangIn =>
            (36, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  format!("{}{}", anim_scale_xml(spid, id_fx, 10000, 100000, 700),
                          anim_fly_xml(spid, id_fx2, "ppt_x", 1.0, 0.0, 700)))),

        // ── Bounce ──────────────────────────────────────────────────────
        // Drops from above with sinusoidal bounce keyframes, revealed by wipe(down)
        AnimationEffectType::BounceIn => {
            let bounce_y = anim_keyframes_xml(spid, id_fx2, "ppt_y", &[
                (0.0, -1.5), (0.365, 0.08), (0.55, -0.04),
                (0.72, 0.015), (0.85, -0.005), (1.0, 0.0),
            ], 1600);
            (26, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  format!("{}{}", anim_effect_filter(spid, id_fx, "in", "wipe(down)", 580), bounce_y)))
        }

        // ── Credits ─────────────────────────────────────────────────────
        // Scrolls up from below (slow rise)
        AnimationEffectType::CreditsIn =>
            (32, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_fly_xml(spid, id_fx, "ppt_y", 1.0, 0.0, 2000))),

        // ── Curve Up ────────────────────────────────────────────────────
        // Scales down from oversized while following a bezier arc into position, with fade
        AnimationEffectType::CurveUpIn => {
            let id_fx3 = id_fx2 + 1;
            let path = "M -0.46736 0.92887 C -0.37517 0.88508 -0.02552 0.75279 \
                        0.0908 0.66613 C 0.20747 0.57948 0.21649 0.50394 \
                        0.23177 0.40825 C 0.24705 0.31256 0.22118 0.15964 \
                        0.18264 0.09152 C 0.1441 0.02341 0.03802 0.0 0.0 0.0";
            (52, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  format!("{}{}{}",
                      anim_scale_decel_xy_xml(spid, id_fx, 250000, 250000, 100000, 100000, 1000, 50000),
                      anim_motion_xml(spid, id_fx2, path, 1000, 50000),
                      anim_effect_filter(spid, id_fx3, "in", "fade", 1000))))
        }

        // ── Drop ────────────────────────────────────────────────────────
        // Falls in from above
        AnimationEffectType::DropIn =>
            (34, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_fly_xml(spid, id_fx, "ppt_y", -1.0, 0.0, 500))),

        // ── Flip ────────────────────────────────────────────────────────
        // Approximated as horizontal swivel (width grows from 0)
        AnimationEffectType::FlipIn =>
            (35, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_scale_xy_xml(spid, id_fx, 0, 100000, 100000, 100000, 500))),

        // ── Pinwheel ────────────────────────────────────────────────────
        // Fast spin + scale (like Spinner but faster rotation)
        AnimationEffectType::PinwheelIn => {
            let rot_by = (720.0_f32 * 60_000.0).round() as i64; // 2 full spins
            (37, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  format!("{}{}", anim_scale_xml(spid, id_fx, 10000, 100000, 600),
                          anim_rot_xml(spid, id_fx2, rot_by, 600))))
        }

        // ── Spiral In ───────────────────────────────────────────────────
        // Approximated as scale + fly from off-screen diagonal
        AnimationEffectType::SpiralIn =>
            (38, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  format!("{}{}", anim_scale_xml(spid, id_fx, 10000, 100000, 700),
                          anim_fly_xml(spid, id_fx2, "ppt_x", -0.5, 0.0, 700)))),

        // ── Basic Swivel ────────────────────────────────────────────────
        // Quarter rotation entrance
        AnimationEffectType::BasicSwivelIn => {
            let rot_by = (90.0_f32 * 60_000.0).round() as i64;
            (39, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  anim_rot_xml(spid, id_fx, rot_by, 500)))
        }

        // ── Whip ────────────────────────────────────────────────────────
        // Fast fly-in from right with scale
        AnimationEffectType::WhipIn =>
            (40, "entr", 0,
             entr(set_visibility_xml(spid, id_set, "visible", 0),
                  format!("{}{}", anim_scale_xml(spid, id_fx, 10000, 100000, 400),
                          anim_fly_xml(spid, id_fx2, "ppt_x", 1.0, 0.0, 400)))),

        // ── Exit Basic (filter-based mirrors of entrance counterparts) ──
        AnimationEffectType::BlindsOut(orient) => {
            let (subtype, filter) = blinds_filter(orient);
            (3, "exit", subtype,
             exit_xml(anim_effect_filter(spid, id_fx, "out", &filter, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }
        AnimationEffectType::CheckerboardOut(dir) => {
            let (subtype, filter) = checkerboard_filter(dir);
            (5, "exit", subtype,
             exit_xml(anim_effect_filter(spid, id_fx, "out", &filter, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }
        AnimationEffectType::DissolveOut =>
            (12, "exit", 0,
             exit_xml(anim_effect_filter(spid, id_fx, "out", "dissolve()", 500),
                      set_visibility_xml(spid, id_set, "hidden", 500))),
        AnimationEffectType::PeekOut(dir) => {
            let (subtype, filter) = wipe_filter(dir);
            (13, "exit", subtype,
             exit_xml(anim_effect_filter(spid, id_fx, "out", &filter, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }
        AnimationEffectType::RandomBarsOut(orient) => {
            let (subtype, filter) = random_bars_filter(orient);
            (14, "exit", subtype,
             exit_xml(anim_effect_filter(spid, id_fx, "out", &filter, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }
        AnimationEffectType::ShapeOut(variant) => {
            let (preset_id, filter) = shape_filter(variant, false);
            (preset_id, "exit", 0,
             exit_xml(anim_effect_filter(spid, id_fx, "out", &filter, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }
        AnimationEffectType::StripsOut(dir) => {
            let (subtype, filter) = strips_filter(dir);
            (6, "exit", subtype,
             exit_xml(anim_effect_filter(spid, id_fx, "out", &filter, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }
        AnimationEffectType::WedgeOut =>
            (17, "exit", 0,
             exit_xml(anim_effect_filter(spid, id_fx, "out", "wedge()", 500),
                      set_visibility_xml(spid, id_set, "hidden", 500))),
        AnimationEffectType::WheelOut(spokes) => {
            let n = (*spokes).max(1);
            let filter = format!("wheel(spokes={n})");
            (18, "exit", n,
             exit_xml(anim_effect_filter(spid, id_fx, "out", &filter, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }

        // ── Exit Subtle ─────────────────────────────────────────────────
        AnimationEffectType::ContractOut =>
            (22, "exit", 0,
             exit_xml(anim_scale_xml(spid, id_fx, 100000, 0, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500))),
        AnimationEffectType::SwivelOut =>
            (21, "exit", 0,
             exit_xml(anim_scale_xy_xml(spid, id_fx, 100000, 100000, 0, 100000, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500))),

        // ── Exit Moderate ───────────────────────────────────────────────
        AnimationEffectType::CentreRevolveOut => {
            let rot_by = (720.0_f32 * 60_000.0).round() as i64;
            (23, "exit", 0,
             exit_xml(
                 format!("{}{}", anim_scale_xml(spid, id_fx, 100000, 10000, 700),
                         anim_rot_xml(spid, id_fx2, rot_by, 700)),
                 set_visibility_xml(spid, id_set, "hidden", 700)))
        }
        AnimationEffectType::CollapseOut =>
            (31, "exit", 0,
             exit_xml(anim_scale_xy_xml(spid, id_fx, 100000, 100000, 100000, 0, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500))),
        AnimationEffectType::FloatOut(_dir) =>
            (30, "exit", 0,
             exit_xml(
                 format!("{}{}", anim_effect_filter(spid, id_fx, "out", "fade", 800),
                         anim_style_rotation_xml(spid, id_fx2, 0.0, 90.0, 800)),
                 set_visibility_xml(spid, id_set, "hidden", 800))),
        AnimationEffectType::ShrinkTurnOut => {
            let rot_by = (90.0_f32 * 60_000.0).round() as i64;
            (24, "exit", 0,
             exit_xml(
                 format!("{}{}", anim_scale_xml(spid, id_fx, 100000, 10000, 500),
                         anim_rot_xml(spid, id_fx2, rot_by, 500)),
                 set_visibility_xml(spid, id_set, "hidden", 500)))
        }
        AnimationEffectType::SinkDownOut =>
            (25, "exit", 0,
             exit_xml(anim_fly_xml(spid, id_fx, "ppt_y", 0.0, 0.5, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500))),
        AnimationEffectType::SpinnerOut => {
            let rot_by = (360.0_f32 * 60_000.0).round() as i64;
            (28, "exit", 0,
             exit_xml(
                 format!("{}{}", anim_scale_xml(spid, id_fx, 100000, 10000, 700),
                         anim_rot_xml(spid, id_fx2, rot_by, 700)),
                 set_visibility_xml(spid, id_set, "hidden", 700)))
        }
        AnimationEffectType::BasicZoomOut =>
            (27, "exit", 0,
             exit_xml(anim_scale_xml(spid, id_fx, 100000, 10000, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500))),
        AnimationEffectType::StretchyOut(dir) => {
            let (to_x, to_y, subtype) = match dir {
                Direction::Left | Direction::Right => (0, 100000, 4u32),
                Direction::Up   | Direction::Down  => (100000, 0, 8u32),
            };
            (29, "exit", subtype,
             exit_xml(anim_scale_xy_xml(spid, id_fx, 100000, 100000, to_x, to_y, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }

        // ── Exit Exciting ───────────────────────────────────────────────
        AnimationEffectType::BoomerangOut =>
            (36, "exit", 0,
             exit_xml(
                 format!("{}{}", anim_scale_xml(spid, id_fx, 100000, 10000, 700),
                         anim_fly_xml(spid, id_fx2, "ppt_x", 0.0, 1.0, 700)),
                 set_visibility_xml(spid, id_set, "hidden", 700))),
        AnimationEffectType::BounceOut => {
            // Brief bounce up, then fall off the bottom of the slide
            let bounce_y = anim_keyframes_xml(spid, id_fx, "ppt_y", &[
                (0.0, 0.0), (0.15, -0.015), (0.30, 0.005),
                (0.45, -0.003), (0.55, 0.0), (1.0, 1.5),
            ], 1600);
            (26, "exit", 0,
             exit_xml(bounce_y, set_visibility_xml(spid, id_set, "hidden", 1600)))
        }
        AnimationEffectType::CreditsOut =>
            (32, "exit", 0,
             exit_xml(anim_fly_xml(spid, id_fx, "ppt_y", 0.0, -1.0, 2000),
                      set_visibility_xml(spid, id_set, "hidden", 2000))),
        AnimationEffectType::CurveDownOut => {
            // Scale up while following a downward bezier arc, with fade out
            let id_fx3 = id_fx2 + 1;
            let path = "M 0.0 0.0 C 0.03802 0.0 0.1441 0.02341 0.18264 0.09152 \
                        C 0.22118 0.15964 0.24705 0.31256 0.23177 0.40825 \
                        C 0.21649 0.50394 0.20747 0.57948 0.0908 0.66613 \
                        C -0.02552 0.75279 -0.37517 0.88508 -0.46736 0.92887";
            (52, "exit", 0,
             exit_xml(
                 format!("{}{}{}",
                     anim_scale_decel_xy_xml(spid, id_fx, 100000, 100000, 250000, 250000, 1000, 50000),
                     anim_motion_xml(spid, id_fx2, path, 1000, 50000),
                     anim_effect_filter(spid, id_fx3, "out", "fade", 1000)),
                 set_visibility_xml(spid, id_set, "hidden", 1000)))
        }
        AnimationEffectType::DropOut =>
            (34, "exit", 0,
             exit_xml(anim_fly_xml(spid, id_fx, "ppt_y", 0.0, 1.0, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500))),
        AnimationEffectType::FlipOut =>
            (35, "exit", 0,
             exit_xml(anim_scale_xy_xml(spid, id_fx, 100000, 100000, 0, 100000, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500))),
        AnimationEffectType::PinwheelOut => {
            let rot_by = (720.0_f32 * 60_000.0).round() as i64;
            (37, "exit", 0,
             exit_xml(
                 format!("{}{}", anim_scale_xml(spid, id_fx, 100000, 10000, 600),
                         anim_rot_xml(spid, id_fx2, rot_by, 600)),
                 set_visibility_xml(spid, id_set, "hidden", 600)))
        }
        AnimationEffectType::SpiralOut =>
            (38, "exit", 0,
             exit_xml(
                 format!("{}{}", anim_scale_xml(spid, id_fx, 100000, 10000, 700),
                         anim_fly_xml(spid, id_fx2, "ppt_x", 0.0, 0.5, 700)),
                 set_visibility_xml(spid, id_set, "hidden", 700))),
        AnimationEffectType::BasicSwivelOut => {
            let rot_by = (90.0_f32 * 60_000.0).round() as i64;
            (39, "exit", 0,
             exit_xml(anim_rot_xml(spid, id_fx, rot_by, 500),
                      set_visibility_xml(spid, id_set, "hidden", 500)))
        }
        AnimationEffectType::WhipOut =>
            (40, "exit", 0,
             exit_xml(
                 format!("{}{}", anim_scale_xml(spid, id_fx, 100000, 10000, 400),
                         anim_fly_xml(spid, id_fx2, "ppt_x", 0.0, 1.0, 400)),
                 set_visibility_xml(spid, id_set, "hidden", 400))),

        // ── Emphasis (Basic) ────────────────────────────────────────────
        AnimationEffectType::Spin(degrees) => {
            let by = (*degrees * 60_000.0).round() as i64;
            (8, "emph", 0, anim_rot_xml(spid, id_fx, by, 2000))
        }
        AnimationEffectType::Pulse =>
            (14, "emph", 0, anim_pulse_xml(spid, id_fx, 500)),
        AnimationEffectType::GrowShrink(scale) => {
            let to_val = (*scale * 100_000.0).round() as u32;
            (18, "emph", 0, anim_scale_xml(spid, id_fx, 100000, to_val, 500))
        }
        AnimationEffectType::FillColor(hex) =>
            (1, "emph", 0, anim_clr_to_xml(spid, id_fx, "fillcolor", hex, 500, false)),
        AnimationEffectType::FontColor(hex) =>
            (2, "emph", 0, anim_clr_to_xml(spid, id_fx, "style.color", hex, 500, false)),
        AnimationEffectType::LineColor(hex) =>
            (3, "emph", 0, anim_clr_to_xml(spid, id_fx, "strokecolor", hex, 500, false)),
        AnimationEffectType::Transparency(level) => {
            let val = level.clamp(0.0, 1.0);
            (4, "emph", 0, anim_opacity_xml(spid, id_fx, val, 500, true))
        }

        // ── Emphasis (Subtle) ────────────────────────────────────────────
        AnimationEffectType::BoldFlash =>
            (5, "emph", 0,
             anim_str_discrete_xml(spid, id_fx, "style.fontWeight", "normal", "bold", 500, true)),
        AnimationEffectType::BrushColor(hex) =>
            (6, "emph", 0, anim_clr_to_xml(spid, id_fx, "fillcolor", hex, 500, true)),
        AnimationEffectType::ComplementaryColor =>
            // Shift fill hue by 180° (10 800 000 = 180° in 1/60000° units)
            (7, "emph", 0,
             anim_clr_hsl_by_xml(spid, id_fx, "fillcolor", 10_800_000, 0, 0, 500, true)),
        AnimationEffectType::ComplementaryColor2 =>
            // Shift fill hue by 120° (7 200 000 = 120° in 1/60000° units)
            (9, "emph", 0,
             anim_clr_hsl_by_xml(spid, id_fx, "fillcolor", 7_200_000, 0, 0, 500, true)),
        AnimationEffectType::ContrastingColor =>
            // Shift luminance by 50 % (moves toward opposite brightness)
            (10, "emph", 0,
             anim_clr_hsl_by_xml(spid, id_fx, "fillcolor", 0, 0, 50_000, 500, true)),
        AnimationEffectType::Darken =>
            // Reduce opacity to 55 % — simulates fill darkening and reverses
            (11, "emph", 0, anim_opacity_xml(spid, id_fx, 0.55, 500, true)),
        AnimationEffectType::Desaturate =>
            // Flash fill toward neutral grey ("808080") and revert
            (12, "emph", 0, anim_clr_to_xml(spid, id_fx, "fillcolor", "808080", 500, true)),
        AnimationEffectType::Lighten =>
            // Flash fill toward near-white ("E8E8E8") and revert
            (13, "emph", 0, anim_clr_to_xml(spid, id_fx, "fillcolor", "E8E8E8", 500, true)),
        AnimationEffectType::ObjectColor(hex) =>
            (15, "emph", 0, anim_clr_to_xml(spid, id_fx, "fillcolor", hex, 500, true)),
        AnimationEffectType::Underline =>
            (16, "emph", 0,
             anim_str_discrete_xml(spid, id_fx, "style.textDecoration", "none", "underline", 500, true)),

        // ── Emphasis (Moderate) ──────────────────────────────────────────
        AnimationEffectType::ColorPulse(hex) =>
            (17, "emph", 0, anim_clr_to_xml(spid, id_fx, "fillcolor", hex, 500, true)),
        AnimationEffectType::GrowWithColor(hex) => {
            // Grow to 115 % (hold) + change fill colour (hold)
            (19, "emph", 0,
             format!("{}{}",
                 anim_scale_xml(spid, id_fx, 100000, 115000, 500),
                 anim_clr_to_xml(spid, id_fx2, "fillcolor", hex, 500, false)))
        }
        AnimationEffectType::Shimmer =>
            // 3 rapid opacity cycles (250 ms each, auto-reverse each)
            (20, "emph", 0, format!("\
<p:anim calcmode=\"lin\" valueType=\"num\">\
<p:cBhvr>\
<p:cTn id=\"{id_fx}\" dur=\"250\" repeatCount=\"3\" autoRev=\"1\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst>\
</p:cBhvr>\
<p:tavLst>\
<p:tav tm=\"0\"><p:val><p:fltVal val=\"1\"/></p:val></p:tav>\
<p:tav tm=\"100000\"><p:val><p:fltVal val=\"0.25\"/></p:val></p:tav>\
</p:tavLst></p:anim>")),
        AnimationEffectType::Teeter =>
            // Rock ±4° with 4 swings, returning to 0°
            (21, "emph", 0, anim_keyframes_xml(spid, id_fx, "style.rotation", &[
                (0.0, 0.0), (0.2, 4.0), (0.4, -4.0), (0.6, 4.0), (0.8, -4.0), (1.0, 0.0),
            ], 700)),

        // ── Emphasis (Exciting) ──────────────────────────────────────────
        AnimationEffectType::Blink =>
            // 3 blink cycles: alternating visible/hidden keyframes
            (22, "emph", 0, format!("\
<p:anim calcmode=\"discrete\" valueType=\"str\">\
<p:cBhvr>\
<p:cTn id=\"{id_fx}\" dur=\"750\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>\
</p:cBhvr>\
<p:tavLst>\
<p:tav tm=\"0\"><p:val><p:strVal val=\"visible\"/></p:val></p:tav>\
<p:tav tm=\"16667\"><p:val><p:strVal val=\"hidden\"/></p:val></p:tav>\
<p:tav tm=\"33333\"><p:val><p:strVal val=\"visible\"/></p:val></p:tav>\
<p:tav tm=\"50000\"><p:val><p:strVal val=\"hidden\"/></p:val></p:tav>\
<p:tav tm=\"66667\"><p:val><p:strVal val=\"visible\"/></p:val></p:tav>\
<p:tav tm=\"83333\"><p:val><p:strVal val=\"hidden\"/></p:val></p:tav>\
<p:tav tm=\"100000\"><p:val><p:strVal val=\"visible\"/></p:val></p:tav>\
</p:tavLst></p:anim>")),
        AnimationEffectType::BoldReveal => {
            // Bold flash + brief scale-up
            let bold  = anim_str_discrete_xml(spid, id_fx,  "style.fontWeight", "normal", "bold", 500, true);
            let scale = anim_scale_xml(spid, id_fx2, 100000, 112000, 250);
            (23, "emph", 0, format!("{bold}{scale}"))
        }
        AnimationEffectType::Wave =>
            // Oscillating rotation: 5 swings with decreasing amplitude
            (24, "emph", 0, anim_keyframes_xml(spid, id_fx, "style.rotation", &[
                (0.0, 0.0),  (0.1, -5.0), (0.2, 5.0),  (0.3, -5.0), (0.4, 5.0),
                (0.5, -3.0), (0.6, 3.0),  (0.7, -2.0), (0.8, 2.0),  (1.0, 0.0),
            ], 1000)),
    }
}

// ── Inner animation XML builders ──────────────────────────────────────────────

/// `<p:set>` that toggles `style.visibility` to `val`, starting at `delay_ms` milliseconds.
/// Pass `delay_ms=0` for entrance effects (fires at animation start).
/// Pass `delay_ms={dur}` for exit effects (fires after animation completes).
fn set_visibility_xml(spid: usize, id: usize, val: &str, delay_ms: u32) -> String {
    format!(
        "<p:set><p:cBhvr>\
<p:cTn id=\"{id}\" dur=\"1\" fill=\"hold\">\
<p:stCondLst><p:cond delay=\"{delay_ms}\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>\
</p:cBhvr><p:to><p:strVal val=\"{val}\"/></p:to></p:set>"
    )
}

/// `<p:animEffect>` with a SMIL filter (Fade, Wipe, Split, …).
fn anim_effect_filter(spid: usize, id: usize, transition: &str, filter: &str, dur: u32) -> String {
    format!(
        "<p:animEffect transition=\"{transition}\" filter=\"{filter}\">\
<p:cBhvr>\
<p:cTn id=\"{id}\" dur=\"{dur}\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
</p:cBhvr></p:animEffect>"
    )
}

/// `<p:animScale>` with independent X and Y axis values (for asymmetric effects like Swivel/Stretch).
fn anim_scale_xy_xml(
    spid: usize, id: usize,
    from_x: u32, from_y: u32,
    to_x: u32, to_y: u32,
    dur: u32,
) -> String {
    format!(
        "<p:animScale>\
<p:cBhvr><p:cTn id=\"{id}\" dur=\"{dur}\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl></p:cBhvr>\
<p:from x=\"{from_x}\" y=\"{from_y}\"/>\
<p:to x=\"{to_x}\" y=\"{to_y}\"/>\
</p:animScale>"
    )
}

/// `<p:animScale>` from `from_val` to `to_val` (values in units of 1/1000 %, e.g. 100000 = 100%).
fn anim_scale_xml(spid: usize, id: usize, from_val: u32, to_val: u32, dur: u32) -> String {
    format!(
        "<p:animScale>\
<p:cBhvr><p:cTn id=\"{id}\" dur=\"{dur}\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl></p:cBhvr>\
<p:from x=\"{from_val}\" y=\"{from_val}\"/>\
<p:to x=\"{to_val}\" y=\"{to_val}\"/>\
</p:animScale>"
    )
}

/// `<p:animScale>` grow to 115 % and auto-reverse back (Pulse).
fn anim_pulse_xml(spid: usize, id: usize, dur: u32) -> String {
    let half = dur / 2;
    format!(
        "<p:animScale>\
<p:cBhvr><p:cTn id=\"{id}\" dur=\"{half}\" autoRev=\"1\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl></p:cBhvr>\
<p:to x=\"115000\" y=\"115000\"/>\
</p:animScale>"
    )
}

/// `<p:anim>` for position-based Fly effects using additive offsets from natural position.
/// `start_val` and `end_val` are fractional slide offsets (0.0 = natural, 1.0 = full slide width/height).
fn anim_fly_xml(
    spid: usize, id: usize,
    attr: &str, start_val: f32, end_val: f32,
    dur: u32,
) -> String {
    format!(
        "<p:anim calcmode=\"lin\" valueType=\"num\">\
<p:cBhvr additive=\"sum\">\
<p:cTn id=\"{id}\" dur=\"{dur}\" fill=\"hold\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst>\
</p:cBhvr>\
<p:tavLst>\
<p:tav tm=\"0\"><p:val><p:fltVal val=\"{start_val}\"/></p:val></p:tav>\
<p:tav tm=\"100000\"><p:val><p:fltVal val=\"{end_val}\"/></p:val></p:tav>\
</p:tavLst></p:anim>"
    )
}

/// `<p:anim>` targeting `style.rotation` — animates CSS rotation from `from_deg` to `to_deg`.
/// Used by Float In (rotates from -90° to 0°).
fn anim_style_rotation_xml(spid: usize, id: usize, from_deg: f32, to_deg: f32, dur: u32) -> String {
    format!(
        "<p:anim calcmode=\"lin\" valueType=\"num\">\
<p:cBhvr>\
<p:cTn id=\"{id}\" dur=\"{dur}\" decel=\"100000\" fill=\"hold\"/>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>style.rotation</p:attrName></p:attrNameLst>\
</p:cBhvr>\
<p:tavLst>\
<p:tav tm=\"0\"><p:val><p:fltVal val=\"{from_deg}\"/></p:val></p:tav>\
<p:tav tm=\"100000\"><p:val><p:fltVal val=\"{to_deg}\"/></p:val></p:tav>\
</p:tavLst></p:anim>"
    )
}

/// `<p:animRot>` — spin by `by_units` (60 000 units per degree), clockwise positive.
fn anim_rot_xml(spid: usize, id: usize, by_units: i64, dur: u32) -> String {
    format!(
        "<p:animRot by=\"{by_units}\">\
<p:cBhvr>\
<p:cTn id=\"{id}\" dur=\"{dur}\" fill=\"hold\"/>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst>\
</p:cBhvr></p:animRot>"
    )
}

/// `<p:animClr>` that changes a colour attribute to a specific RGB target.
/// `attr`: `"fillcolor"` | `"strokecolor"` | `"style.color"`.
/// `auto_rev=true` pulses back to the original after reaching the target.
fn anim_clr_to_xml(spid: usize, id: usize, attr: &str, hex: &str, dur: u32, auto_rev: bool) -> String {
    let ar = if auto_rev { " autoRev=\"1\"" } else { "" };
    // <p:to> in CT_TLAnimateColorBehavior takes a bare color element (not CT_TLAnimVariant)
    format!("<p:animClr clrSpc=\"rgb\" dir=\"cw\">\
<p:cBhvr>\
<p:cTn id=\"{id}\" dur=\"{dur}\"{ar} fill=\"hold\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst>\
</p:cBhvr>\
<p:to><a:srgbClr val=\"{hex}\"/></p:to>\
</p:animClr>")
}

/// `<p:animClr>` with an HSL relative `<p:by>` adjustment.
/// Used by ComplementaryColor (hue shift), ContrastingColor (lum shift), etc.
/// Units — h: 0..21 600 000 (0°..360° in 1/60000° steps);
///          s/l: signed 1/1000% units (50 000 = +50%, -50 000 = -50%).
fn anim_clr_hsl_by_xml(spid: usize, id: usize, attr: &str,
                        h: i64, s: i64, l: i64,
                        dur: u32, auto_rev: bool) -> String {
    let ar = if auto_rev { " autoRev=\"1\"" } else { "" };
    // <p:by> uses <p:hsl h="..." s="..." l="..."/> for relative HSL adjustments
    format!("<p:animClr clrSpc=\"hsl\" dir=\"cw\">\
<p:cBhvr>\
<p:cTn id=\"{id}\" dur=\"{dur}\"{ar} fill=\"hold\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst>\
</p:cBhvr>\
<p:by><p:hsl h=\"{h}\" s=\"{s}\" l=\"{l}\"/></p:by>\
</p:animClr>")
}

/// `<p:anim>` targeting `style.opacity`.  Goes from 1.0 → `target` (0.0–1.0);
/// `auto_rev=true` bounces back to 1.0 (used for Transparency, Darken, Shimmer).
fn anim_opacity_xml(spid: usize, id: usize, target: f32, dur: u32, auto_rev: bool) -> String {
    let ar = if auto_rev { " autoRev=\"1\"" } else { "" };
    format!("<p:anim calcmode=\"lin\" valueType=\"num\">\
<p:cBhvr>\
<p:cTn id=\"{id}\" dur=\"{dur}\"{ar}>\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst>\
</p:cBhvr>\
<p:tavLst>\
<p:tav tm=\"0\"><p:val><p:fltVal val=\"1\"/></p:val></p:tav>\
<p:tav tm=\"100000\"><p:val><p:fltVal val=\"{target}\"/></p:val></p:tav>\
</p:tavLst></p:anim>")
}

/// Discrete string-valued `<p:anim>` for style properties (fontWeight, textDecoration, …).
/// `auto_rev=true` plays back from `to_val` to `from_val` after reaching the target.
fn anim_str_discrete_xml(spid: usize, id: usize, attr: &str,
                          from_val: &str, to_val: &str,
                          dur: u32, auto_rev: bool) -> String {
    let ar = if auto_rev { " autoRev=\"1\"" } else { "" };
    format!("<p:anim calcmode=\"discrete\" valueType=\"str\">\
<p:cBhvr>\
<p:cTn id=\"{id}\" dur=\"{dur}\"{ar}>\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst>\
</p:cBhvr>\
<p:tavLst>\
<p:tav tm=\"0\"><p:val><p:strVal val=\"{from_val}\"/></p:val></p:tav>\
<p:tav tm=\"100000\"><p:val><p:strVal val=\"{to_val}\"/></p:val></p:tav>\
</p:tavLst></p:anim>")
}

/// Multi-keyframe position animation (for Bounce effects).
/// `keyframes`: pairs of `(time_0_to_1, offset_fraction)`.
fn anim_keyframes_xml(spid: usize, id: usize, attr: &str, keyframes: &[(f32, f32)], dur: u32) -> String {
    let tav_list: String = keyframes.iter().map(|(t, v)| {
        let tm = (*t * 100_000.0).round() as u32;
        format!("<p:tav tm=\"{tm}\"><p:val><p:fltVal val=\"{v}\"/></p:val></p:tav>")
    }).collect();
    format!("<p:anim calcmode=\"lin\" valueType=\"num\">\
<p:cBhvr additive=\"sum\">\
<p:cTn id=\"{id}\" dur=\"{dur}\" fill=\"hold\">\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
<p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst>\
</p:cBhvr>\
<p:tavLst>{tav_list}</p:tavLst>\
</p:anim>")
}

/// `<p:animMotion>` with an SVG-like bezier path (origin=layout, pathEditMode=relative).
/// Used by Curve Up / Curve Down.
fn anim_motion_xml(spid: usize, id: usize, path: &str, dur: u32, decelerate: u32) -> String {
    let decel_attr = if decelerate > 0 { format!(" decelerate=\"{decelerate}\"") } else { String::new() };
    format!("<p:animMotion origin=\"layout\" path=\"{path}\" pathEditMode=\"relative\">\
<p:cBhvr>\
<p:cTn id=\"{id}\" dur=\"{dur}\"{decel_attr}>\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
</p:cBhvr>\
</p:animMotion>")
}

/// `<p:animScale>` with independent X/Y axes and an optional `decelerate` attribute.
/// Used by Curve Up / Curve Down where the scale starts oversized (250 %) and shrinks.
fn anim_scale_decel_xy_xml(
    spid: usize, id: usize,
    from_x: u32, from_y: u32,
    to_x: u32, to_y: u32,
    dur: u32, decelerate: u32,
) -> String {
    let decel_attr = if decelerate > 0 { format!(" decelerate=\"{decelerate}\"") } else { String::new() };
    format!("<p:animScale>\
<p:cBhvr>\
<p:cTn id=\"{id}\" dur=\"{dur}\"{decel_attr}>\
<p:stCondLst><p:cond delay=\"0\"/></p:stCondLst></p:cTn>\
<p:tgtEl><p:spTgt spid=\"{spid}\"/></p:tgtEl>\
</p:cBhvr>\
<p:from x=\"{from_x}\" y=\"{from_y}\"/>\
<p:to x=\"{to_x}\" y=\"{to_y}\"/>\
</p:animScale>")
}

// ── Parameter helpers ─────────────────────────────────────────────────────────

/// Returns `(presetSubtype, filter_string)` for Blinds.
fn blinds_filter(orient: &SplitOrientation) -> (u32, String) {
    match orient {
        SplitOrientation::Horizontal => (10, "blinds(horizontal)".to_string()),
        SplitOrientation::Vertical   => (4,  "blinds(vertical)".to_string()),
    }
}

/// Returns `(presetSubtype, filter_string)` for Checkerboard.
fn checkerboard_filter(dir: &CheckerboardDir) -> (u32, String) {
    match dir {
        CheckerboardDir::Across => (10, "checkerboard(across)".to_string()),
        CheckerboardDir::Down   => (4,  "checkerboard(down)".to_string()),
    }
}

/// Returns `(presetSubtype, filter_string)` for Random Bars.
fn random_bars_filter(orient: &SplitOrientation) -> (u32, String) {
    match orient {
        SplitOrientation::Horizontal => (10, "randombar(horizontal)".to_string()),
        SplitOrientation::Vertical   => (4,  "randombar(vertical)".to_string()),
    }
}

/// Returns `(presetID, filter_string)` for Shape entrance/exit.
fn shape_filter(variant: &ShapeVariant, entering: bool) -> (u32, String) {
    let dir = if entering { "in" } else { "out" };
    match variant {
        ShapeVariant::Box     => (8,  format!("box({dir})")),
        ShapeVariant::Circle  => (14, format!("circle({dir})")),
        ShapeVariant::Diamond => (15, format!("diamond({dir})")),
        ShapeVariant::Plus    => (16, format!("plus({dir})")),
    }
}

/// Returns `(presetSubtype, filter_string)` for Strips.
fn strips_filter(dir: &StripDir) -> (u32, String) {
    match dir {
        StripDir::LeftDown  => (9,  "strips(leftdown)".to_string()),
        StripDir::LeftUp    => (3,  "strips(leftup)".to_string()),
        StripDir::RightDown => (12, "strips(rightdown)".to_string()),
        StripDir::RightUp   => (6,  "strips(rightup)".to_string()),
    }
}

/// Returns `(presetSubtype, filter_string)` for Wipe.
fn wipe_filter(dir: &Direction) -> (u32, String) {
    let (subtype, name) = match dir {
        Direction::Left  => (12, "wipe(left)"),
        Direction::Right => (4,  "wipe(right)"),
        Direction::Up    => (8,  "wipe(up)"),
        Direction::Down  => (0,  "wipe(down)"),
    };
    (subtype, name.to_string())
}

/// Returns `(presetSubtype, filter_string)` for Split.
fn split_filter(orient: &SplitOrientation, entering: bool) -> (u32, String) {
    let direction = if entering { "in" } else { "out" };
    let (subtype, orientation) = match orient {
        SplitOrientation::Horizontal => (22, "horizontal"),
        SplitOrientation::Vertical   => (24, "vertical"),
    };
    (subtype, format!("barn(direction={direction},orientation={orientation})"))
}

/// Returns `(presetSubtype, attr_name, start_offset, end_offset)` for Fly.
/// Offsets are fractional slide units added to the natural position (`additive="sum"`).
/// `entering=true` → object starts off-screen and moves to natural position (end=0.0).
/// `entering=false` → object moves from natural position off-screen (start=0.0).
fn fly_params(dir: &Direction, entering: bool) -> (u32, &'static str, f32, f32) {
    match (dir, entering) {
        (Direction::Left,  true)  => (4, "ppt_x", -1.0,  0.0),
        (Direction::Right, true)  => (2, "ppt_x",  1.0,  0.0),
        (Direction::Up,    true)  => (4, "ppt_y", -1.0,  0.0),
        (Direction::Down,  true)  => (8, "ppt_y",  1.0,  0.0),
        (Direction::Left,  false) => (4, "ppt_x",  0.0, -1.0),
        (Direction::Right, false) => (2, "ppt_x",  0.0,  1.0),
        (Direction::Up,    false) => (4, "ppt_y",  0.0, -1.0),
        (Direction::Down,  false) => (8, "ppt_y",  0.0,  1.0),
    }
}
