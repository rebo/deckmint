use crate::enums::ShapeType;
use crate::error::PptxError;
use crate::objects::{SlideObject, SlideRel, SlideRelMedia, TextObject, ShapeObject, ImageObject, ConnectorObject, ConnectorType, ConnectorOptions};
use crate::objects::text::{TextRun, TextRunOptions};
use crate::objects::shape::ShapeOptions;
use crate::objects::image::{ImageOptions, ImageOptionsBuilder};
use crate::packaging::assemble_pptx;
use crate::slide::Slide;
use crate::types::PresLayout;

/// A named slide layout — the anatomy tier between the master and a slide.
///
/// Chrome that's shared by every slide of one pedagogical archetype (a
/// badge/pill, guide-rail outline boxes, an archetype-specific icon) belongs
/// here, drawn ONCE per layout, rather than re-stamped onto every slide that
/// uses it — the master carries chrome common to *every* layout (see
/// [`SlideMasterDef`]), a layout carries chrome common to every slide bound to
/// *it* (via [`Presentation::add_slide_with_layout`]), and a slide should carry
/// only its own content.
///
/// Use [`Presentation::define_layout`] to create one and get back a `&mut
/// SlideLayout` to build on, then add objects with the same convenience
/// methods [`SlideMasterDef`] offers: [`add_text`](SlideLayout::add_text),
/// [`add_shape`](SlideLayout::add_shape), [`add_image`](SlideLayout::add_image).
#[derive(Debug, Clone, Default)]
pub struct SlideLayout {
    pub name: String,
    /// Optional custom width in inches (defaults to presentation width)
    pub width: Option<f64>,
    /// Optional custom height in inches (defaults to presentation height)
    pub height: Option<f64>,
    /// Objects to render on this layout (chrome shapes/text/images).
    pub objects: Vec<SlideObject>,
    /// Non-media relationships (hyperlinks, etc.)
    pub(crate) rels: Vec<SlideRel>,
    /// Media relationships (images, etc.)
    pub(crate) rels_media: Vec<SlideRelMedia>,
    /// Next object counter for auto-naming
    pub(crate) obj_counter: u32,
}

impl SlideLayout {
    fn next_obj_name(&mut self, prefix: &str) -> String {
        self.obj_counter += 1;
        format!("{} {}", prefix, self.obj_counter)
    }

    /// Layout rIds are rebased at packaging time (rId1 is always reserved for
    /// this layout's slideMaster relationship); internally count from 1, same
    /// convention as [`SlideMasterDef::allocate_rid`].
    fn allocate_rid(&self) -> u32 {
        (self.rels.len() + self.rels_media.len() + 1) as u32
    }

    fn register_hyperlink_rel(&mut self, hl: &mut crate::types::HyperlinkProps) {
        if hl.action.is_none() {
            let rid = self.allocate_rid();
            hl.r_id = rid;
            let (rel_type, target, data_field) = if let Some(slide) = hl.slide {
                ("hyperlink".to_string(), slide.to_string(), Some("slide".to_string()))
            } else if let Some(ref url) = hl.url {
                ("hyperlink".to_string(), url.clone(), None)
            } else {
                ("hyperlink".to_string(), String::new(), None)
            };
            self.rels.push(SlideRel { r_id: rid, rel_type, target, data: data_field });
        }
    }

    // ─── Text ───────────────────────────────────────────────

    /// Add a simple text box.
    pub fn add_text(&mut self, text: impl Into<String>, opts: crate::objects::text::TextOptions) -> &mut Self {
        let name = self.next_obj_name("TextBox");
        let run = TextRun { text: text.into(), options: TextRunOptions::default(), break_line: false, soft_break_before: false, field: None, equation_omml: None };
        let obj = TextObject { object_name: name, text: vec![run], options: opts };
        self.objects.push(SlideObject::Text(obj));
        self
    }

    /// Add a text box from multiple TextRun segments (rich text).
    pub fn add_text_runs(&mut self, runs: Vec<TextRun>, opts: crate::objects::text::TextOptions) -> &mut Self {
        let name = self.next_obj_name("TextBox");
        let mut runs = runs;
        for run in &mut runs {
            if let Some(ref mut hl) = run.options.hyperlink {
                self.register_hyperlink_rel(hl);
            }
        }
        let obj = TextObject { object_name: name, text: runs, options: opts };
        self.objects.push(SlideObject::Text(obj));
        self
    }

    // ─── Shape ──────────────────────────────────────────────

    /// Add a shape.
    pub fn add_shape(&mut self, shape_type: ShapeType, mut opts: ShapeOptions) -> &mut Self {
        if let Some(ref mut hl) = opts.hyperlink {
            self.register_hyperlink_rel(hl);
        }
        if let Some(ref mut hl) = opts.hover {
            self.register_hyperlink_rel(hl);
        }
        let name = self.next_obj_name("Shape");
        let obj = ShapeObject { object_name: name, shape_type, options: opts, text: None };
        self.objects.push(SlideObject::Shape(obj));
        self
    }

    /// Add a shape with text inside it.
    pub fn add_shape_with_text(
        &mut self,
        shape_type: ShapeType,
        opts: ShapeOptions,
        text: impl Into<String>,
        text_opts: crate::objects::text::TextOptions,
    ) -> &mut Self {
        let shape_name = self.next_obj_name("Shape");
        let text_name = format!("{} Text", shape_name);
        let run = TextRun { text: text.into(), options: TextRunOptions::default(), break_line: false, soft_break_before: false, field: None, equation_omml: None };
        let text_obj = TextObject { object_name: text_name, text: vec![run], options: text_opts };
        let obj = ShapeObject { object_name: shape_name, shape_type, options: opts, text: Some(text_obj) };
        self.objects.push(SlideObject::Shape(obj));
        self
    }

    // ─── Image ──────────────────────────────────────────────

    /// Add an image from raw bytes.
    pub fn add_image(&mut self, data: Vec<u8>, extension: &str, mut opts: ImageOptions) -> &mut Self {
        let name = self.next_obj_name("Image");
        crate::objects::image::resolve_image_size(&data, &mut opts.position);
        if let Some(ref mut hl) = opts.hyperlink {
            self.register_hyperlink_rel(hl);
        }
        if let Some(ref mut hl) = opts.hover {
            self.register_hyperlink_rel(hl);
        }
        let rid = self.allocate_rid();
        let extn = extension.to_lowercase();
        let is_svg = extn == "svg";
        let target = format!("../media/image{}.{}", rid, extn);
        self.rels_media.push(SlideRelMedia {
            r_id: rid, rel_type: "image".to_string(), target, extn: extn.clone(), data: data.clone(),
        });
        let obj = ImageObject { object_name: name, image_rid: rid, extension: extn, data, is_svg, raster_fallback_rid: None, options: opts };
        self.objects.push(SlideObject::Image(obj));
        self
    }

    /// Add an image from a base64-encoded string.
    pub fn add_image_base64(&mut self, b64: &str, extension: &str, opts: ImageOptions) -> Result<&mut Self, PptxError> {
        let data = crate::utils::decode_base64_image(b64)?;
        Ok(self.add_image(data, extension, opts))
    }

    /// Add an image built via [`ImageOptionsBuilder`] (bytes or base64, whichever the
    /// builder was given) — mirrors [`Slide::add_image_from`](crate::slide::Slide::add_image_from).
    pub fn add_image_from(&mut self, builder: ImageOptionsBuilder) -> Result<&mut Self, PptxError> {
        let (opts, data, data_b64, ext) = builder.build();
        if let Some(b64_str) = data_b64 {
            let extension = ext.unwrap_or_else(|| "png".to_string());
            return self.add_image_base64(&b64_str, &extension, opts);
        }
        if data.is_empty() {
            return Err(PptxError::InvalidArgument(
                "ImageOptionsBuilder has no image data; call .bytes() or .base64() first".to_string(),
            ));
        }
        let extension = ext.unwrap_or_else(|| "png".to_string());
        Ok(self.add_image(data, &extension, opts))
    }

    /// Add an SVG image with a raster (PNG) fallback — mirrors
    /// [`Slide::add_image_svg`](crate::slide::Slide::add_image_svg); see there
    /// for why both media parts are required and independently registered.
    pub fn add_image_svg(&mut self, svg_data: Vec<u8>, png_fallback_data: Vec<u8>, mut opts: ImageOptions) -> &mut Self {
        let name = self.next_obj_name("Image");
        crate::objects::image::resolve_image_size(&png_fallback_data, &mut opts.position);
        if let Some(ref mut hl) = opts.hyperlink {
            self.register_hyperlink_rel(hl);
        }
        if let Some(ref mut hl) = opts.hover {
            self.register_hyperlink_rel(hl);
        }
        let raster_rid = self.allocate_rid();
        self.rels_media.push(SlideRelMedia {
            r_id: raster_rid,
            rel_type: "image".to_string(),
            target: format!("../media/image{raster_rid}.png"),
            extn: "png".to_string(),
            data: png_fallback_data,
        });
        let svg_rid = self.allocate_rid();
        self.rels_media.push(SlideRelMedia {
            r_id: svg_rid,
            rel_type: "image".to_string(),
            target: format!("../media/image{svg_rid}.svg"),
            extn: "svg".to_string(),
            data: svg_data.clone(),
        });
        let obj = ImageObject {
            object_name: name,
            image_rid: svg_rid,
            extension: "svg".to_string(),
            data: svg_data,
            is_svg: true,
            raster_fallback_rid: Some(raster_rid),
            options: opts,
        };
        self.objects.push(SlideObject::Image(obj));
        self
    }

    /// Add an SVG image with a raster (PNG) fallback, both as base64-encoded
    /// strings. See [`Slide::add_image_svg`](crate::slide::Slide::add_image_svg).
    pub fn add_image_svg_base64(&mut self, svg_b64: &str, png_fallback_b64: &str, opts: ImageOptions) -> Result<&mut Self, PptxError> {
        let svg_data = crate::utils::decode_base64_image(svg_b64)?;
        let fallback_data = crate::utils::decode_base64_image(png_fallback_b64)?;
        Ok(self.add_image_svg(svg_data, fallback_data, opts))
    }

    /// Add an SVG image built via [`ImageOptionsBuilder`] — mirrors
    /// [`Slide::add_image_svg_from`](crate::slide::Slide::add_image_svg_from).
    pub fn add_image_svg_from(&mut self, builder: ImageOptionsBuilder) -> Result<&mut Self, PptxError> {
        let (opts, svg_data, svg_b64, fallback_data, fallback_b64) = builder.build_svg();
        let svg_bytes = match svg_b64 {
            Some(b64) => crate::utils::decode_base64_image(&b64)?,
            None => svg_data,
        };
        if svg_bytes.is_empty() {
            return Err(PptxError::InvalidArgument(
                "add_image_svg_from: no SVG data; call .bytes(svg_bytes, \"svg\") or .base64(svg_b64, \"svg\") first".to_string(),
            ));
        }
        let fallback_bytes = match fallback_b64 {
            Some(b64) => crate::utils::decode_base64_image(&b64)?,
            None => fallback_data,
        };
        if fallback_bytes.is_empty() {
            return Err(PptxError::InvalidArgument(
                "add_image_svg_from: no raster fallback; call .svg_fallback(png_bytes) or .svg_fallback_base64(...) — PowerPoint's base <a:blip> requires a real raster image, not just the SVG".to_string(),
            ));
        }
        Ok(self.add_image_svg(svg_bytes, fallback_bytes, opts))
    }

    // ─── Connector ──────────────────────────────────────────

    /// Add a connector line.
    pub fn add_connector(&mut self, connector_type: ConnectorType, opts: ConnectorOptions) -> &mut Self {
        let name = self.next_obj_name("Connector");
        let obj = ConnectorObject { object_name: name, connector_type, options: opts };
        self.objects.push(SlideObject::Connector(obj));
        self
    }

    // ─── Group ──────────────────────────────────────────────

    /// Add a group of shapes.
    pub fn add_group(&mut self, children: Vec<SlideObject>, opts: crate::objects::group::GroupOptions) -> &mut Self {
        use crate::enums::EMU;
        let name = self.next_obj_name("Group");
        let cx = opts.position.w.as_ref().map(|c| c.to_emu(12_192_000)).unwrap_or(EMU as i64);
        let cy = opts.position.h.as_ref().map(|c| c.to_emu(6_858_000)).unwrap_or(EMU as i64);
        let child_extent = opts.child_extent.unwrap_or((cx, cy));
        let obj = crate::objects::GroupObject {
            object_name: name,
            position: opts.position,
            child_offset: opts.child_offset,
            child_extent,
            children,
        };
        self.objects.push(SlideObject::Group(obj));
        self
    }

    // ─── Table ──────────────────────────────────────────────

    /// Add a table.
    pub fn add_table(&mut self, rows: Vec<crate::objects::table::TableRow>, opts: crate::objects::table::TableOptions) -> &mut Self {
        let name = self.next_obj_name("Table");
        let mut opts = opts;
        if opts.object_name.is_none() {
            opts.object_name = Some(name.clone());
        }
        let mut rows = rows;
        for row in &mut rows {
            for cell in row.iter_mut() {
                if let Some(ref mut hl) = cell.options.hyperlink {
                    if hl.action.is_none() {
                        self.register_hyperlink_rel(hl);
                    }
                }
            }
        }
        let obj = crate::objects::TableObject { object_name: name, rows, options: opts };
        self.objects.push(SlideObject::Table(obj));
        self
    }
}

/// Definition of the slide master.
///
/// Use [`SlideMasterDef::new`] to create a master, then add objects with
/// convenience methods like [`add_text`](SlideMasterDef::add_text),
/// [`add_shape`](SlideMasterDef::add_shape), and
/// [`add_image`](SlideMasterDef::add_image).
///
/// Alternatively, build a slide with [`Slide`] methods and then call
/// [`Slide::into_master`] to promote it.
#[derive(Debug, Clone, Default)]
pub struct SlideMasterDef {
    pub title: String,
    /// Background fill color as hex (no #)
    pub background_color: Option<String>,
    /// Background color transparency 0–100 (0 = opaque, 100 = fully transparent)
    pub background_transparency: Option<f64>,
    /// Background image relationship ID (set via [`set_background_image`](SlideMasterDef::set_background_image))
    pub(crate) background_image_rid: Option<u32>,
    /// Objects to render on the master slide
    pub objects: Vec<SlideObject>,
    /// Non-media relationships (hyperlinks, etc.)
    pub(crate) rels: Vec<SlideRel>,
    /// Media relationships (images, etc.)
    pub(crate) rels_media: Vec<SlideRelMedia>,
    /// Next object counter for auto-naming
    pub(crate) obj_counter: u32,
}

impl SlideMasterDef {
    /// Create a new slide master with the given title.
    pub fn new(title: impl Into<String>) -> Self {
        SlideMasterDef {
            title: title.into(),
            ..Default::default()
        }
    }

    fn next_obj_name(&mut self, prefix: &str) -> String {
        self.obj_counter += 1;
        format!("{} {}", prefix, self.obj_counter)
    }

    fn allocate_rid(&self) -> u32 {
        // Master rIds are rebased at packaging time; internally count from 1.
        (self.rels.len() + self.rels_media.len() + 1) as u32
    }

    fn register_hyperlink_rel(&mut self, hl: &mut crate::types::HyperlinkProps) {
        if hl.action.is_none() {
            let rid = self.allocate_rid();
            hl.r_id = rid;
            let (rel_type, target, data_field) = if let Some(slide) = hl.slide {
                ("hyperlink".to_string(), slide.to_string(), Some("slide".to_string()))
            } else if let Some(ref url) = hl.url {
                ("hyperlink".to_string(), url.clone(), None)
            } else {
                ("hyperlink".to_string(), String::new(), None)
            };
            self.rels.push(SlideRel { r_id: rid, rel_type, target, data: data_field });
        }
    }

    // ─── Background ─────────────────────────────────────────

    /// Set the background color as 6-digit hex, no `#` prefix.
    pub fn set_background_color(&mut self, color: impl Into<String>) -> &mut Self {
        self.background_color = Some(color.into().trim_start_matches('#').to_uppercase());
        self
    }

    /// Set background color with transparency (0 = opaque, 100 = fully transparent).
    pub fn set_background_color_transparency(&mut self, color: impl Into<String>, transparency: f64) -> &mut Self {
        self.background_color = Some(color.into().trim_start_matches('#').to_uppercase());
        self.background_transparency = Some(transparency);
        self
    }

    /// Set a background image from raw bytes.
    pub fn set_background_image(&mut self, data: Vec<u8>, extension: &str) -> &mut Self {
        let rid = self.allocate_rid();
        let extn = extension.to_lowercase();
        let target = format!("../media/image{}.{}", rid, extn);
        self.rels_media.push(SlideRelMedia {
            r_id: rid,
            rel_type: "image".to_string(),
            target,
            extn,
            data,
        });
        self.background_image_rid = Some(rid);
        self
    }

    // ─── Text ───────────────────────────────────────────────

    /// Add a simple text box.
    pub fn add_text(&mut self, text: impl Into<String>, opts: crate::objects::text::TextOptions) -> &mut Self {
        let name = self.next_obj_name("TextBox");
        let run = TextRun { text: text.into(), options: TextRunOptions::default(), break_line: false, soft_break_before: false, field: None, equation_omml: None };
        let obj = TextObject { object_name: name, text: vec![run], options: opts };
        self.objects.push(SlideObject::Text(obj));
        self
    }

    /// Add a text box from multiple TextRun segments (rich text).
    pub fn add_text_runs(&mut self, runs: Vec<TextRun>, opts: crate::objects::text::TextOptions) -> &mut Self {
        let name = self.next_obj_name("TextBox");
        let mut runs = runs;
        for run in &mut runs {
            if let Some(ref mut hl) = run.options.hyperlink {
                self.register_hyperlink_rel(hl);
            }
        }
        let obj = TextObject { object_name: name, text: runs, options: opts };
        self.objects.push(SlideObject::Text(obj));
        self
    }

    // ─── Shape ──────────────────────────────────────────────

    /// Add a shape.
    pub fn add_shape(&mut self, shape_type: ShapeType, mut opts: ShapeOptions) -> &mut Self {
        if let Some(ref mut hl) = opts.hyperlink {
            self.register_hyperlink_rel(hl);
        }
        if let Some(ref mut hl) = opts.hover {
            self.register_hyperlink_rel(hl);
        }
        let name = self.next_obj_name("Shape");
        let obj = ShapeObject { object_name: name, shape_type, options: opts, text: None };
        self.objects.push(SlideObject::Shape(obj));
        self
    }

    /// Add a shape with text inside it.
    pub fn add_shape_with_text(
        &mut self,
        shape_type: ShapeType,
        opts: ShapeOptions,
        text: impl Into<String>,
        text_opts: crate::objects::text::TextOptions,
    ) -> &mut Self {
        let shape_name = self.next_obj_name("Shape");
        let text_name = format!("{} Text", shape_name);
        let run = TextRun { text: text.into(), options: TextRunOptions::default(), break_line: false, soft_break_before: false, field: None, equation_omml: None };
        let text_obj = TextObject { object_name: text_name, text: vec![run], options: text_opts };
        let obj = ShapeObject { object_name: shape_name, shape_type, options: opts, text: Some(text_obj) };
        self.objects.push(SlideObject::Shape(obj));
        self
    }

    // ─── Image ──────────────────────────────────────────────

    /// Add an image from raw bytes.
    pub fn add_image(&mut self, data: Vec<u8>, extension: &str, mut opts: ImageOptions) -> &mut Self {
        let name = self.next_obj_name("Image");
        crate::objects::image::resolve_image_size(&data, &mut opts.position);
        if let Some(ref mut hl) = opts.hyperlink {
            self.register_hyperlink_rel(hl);
        }
        if let Some(ref mut hl) = opts.hover {
            self.register_hyperlink_rel(hl);
        }
        let rid = self.allocate_rid();
        let extn = extension.to_lowercase();
        let is_svg = extn == "svg";
        let target = format!("../media/image{}.{}", rid, extn);
        self.rels_media.push(SlideRelMedia {
            r_id: rid, rel_type: "image".to_string(), target, extn: extn.clone(), data: data.clone(),
        });
        let obj = ImageObject { object_name: name, image_rid: rid, extension: extn, data, is_svg, raster_fallback_rid: None, options: opts };
        self.objects.push(SlideObject::Image(obj));
        self
    }

    /// Add an image from a base64-encoded string.
    pub fn add_image_base64(&mut self, b64: &str, extension: &str, opts: ImageOptions) -> Result<&mut Self, PptxError> {
        let data = crate::utils::decode_base64_image(b64)?;
        Ok(self.add_image(data, extension, opts))
    }

    /// Add an image built via [`ImageOptionsBuilder`] (bytes or base64, whichever the
    /// builder was given) — mirrors [`Slide::add_image_from`](crate::slide::Slide::add_image_from).
    pub fn add_image_from(&mut self, builder: ImageOptionsBuilder) -> Result<&mut Self, PptxError> {
        let (opts, data, data_b64, ext) = builder.build();
        if let Some(b64_str) = data_b64 {
            let extension = ext.unwrap_or_else(|| "png".to_string());
            return self.add_image_base64(&b64_str, &extension, opts);
        }
        if data.is_empty() {
            return Err(PptxError::InvalidArgument(
                "ImageOptionsBuilder has no image data; call .bytes() or .base64() first".to_string(),
            ));
        }
        let extension = ext.unwrap_or_else(|| "png".to_string());
        Ok(self.add_image(data, &extension, opts))
    }

    /// Add an SVG image with a raster (PNG) fallback — mirrors
    /// [`Slide::add_image_svg`](crate::slide::Slide::add_image_svg); see there
    /// for why both media parts are required and independently registered.
    pub fn add_image_svg(&mut self, svg_data: Vec<u8>, png_fallback_data: Vec<u8>, mut opts: ImageOptions) -> &mut Self {
        let name = self.next_obj_name("Image");
        crate::objects::image::resolve_image_size(&png_fallback_data, &mut opts.position);
        if let Some(ref mut hl) = opts.hyperlink {
            self.register_hyperlink_rel(hl);
        }
        if let Some(ref mut hl) = opts.hover {
            self.register_hyperlink_rel(hl);
        }
        let raster_rid = self.allocate_rid();
        self.rels_media.push(SlideRelMedia {
            r_id: raster_rid,
            rel_type: "image".to_string(),
            target: format!("../media/image{raster_rid}.png"),
            extn: "png".to_string(),
            data: png_fallback_data,
        });
        let svg_rid = self.allocate_rid();
        self.rels_media.push(SlideRelMedia {
            r_id: svg_rid,
            rel_type: "image".to_string(),
            target: format!("../media/image{svg_rid}.svg"),
            extn: "svg".to_string(),
            data: svg_data.clone(),
        });
        let obj = ImageObject {
            object_name: name,
            image_rid: svg_rid,
            extension: "svg".to_string(),
            data: svg_data,
            is_svg: true,
            raster_fallback_rid: Some(raster_rid),
            options: opts,
        };
        self.objects.push(SlideObject::Image(obj));
        self
    }

    /// Add an SVG image with a raster (PNG) fallback, both as base64-encoded
    /// strings. See [`Slide::add_image_svg`](crate::slide::Slide::add_image_svg).
    pub fn add_image_svg_base64(&mut self, svg_b64: &str, png_fallback_b64: &str, opts: ImageOptions) -> Result<&mut Self, PptxError> {
        let svg_data = crate::utils::decode_base64_image(svg_b64)?;
        let fallback_data = crate::utils::decode_base64_image(png_fallback_b64)?;
        Ok(self.add_image_svg(svg_data, fallback_data, opts))
    }

    /// Add an SVG image built via [`ImageOptionsBuilder`] — mirrors
    /// [`Slide::add_image_svg_from`](crate::slide::Slide::add_image_svg_from).
    pub fn add_image_svg_from(&mut self, builder: ImageOptionsBuilder) -> Result<&mut Self, PptxError> {
        let (opts, svg_data, svg_b64, fallback_data, fallback_b64) = builder.build_svg();
        let svg_bytes = match svg_b64 {
            Some(b64) => crate::utils::decode_base64_image(&b64)?,
            None => svg_data,
        };
        if svg_bytes.is_empty() {
            return Err(PptxError::InvalidArgument(
                "add_image_svg_from: no SVG data; call .bytes(svg_bytes, \"svg\") or .base64(svg_b64, \"svg\") first".to_string(),
            ));
        }
        let fallback_bytes = match fallback_b64 {
            Some(b64) => crate::utils::decode_base64_image(&b64)?,
            None => fallback_data,
        };
        if fallback_bytes.is_empty() {
            return Err(PptxError::InvalidArgument(
                "add_image_svg_from: no raster fallback; call .svg_fallback(png_bytes) or .svg_fallback_base64(...) — PowerPoint's base <a:blip> requires a real raster image, not just the SVG".to_string(),
            ));
        }
        Ok(self.add_image_svg(svg_bytes, fallback_bytes, opts))
    }

    // ─── Connector ──────────────────────────────────────────

    /// Add a connector line.
    pub fn add_connector(&mut self, connector_type: ConnectorType, opts: ConnectorOptions) -> &mut Self {
        let name = self.next_obj_name("Connector");
        let obj = ConnectorObject { object_name: name, connector_type, options: opts };
        self.objects.push(SlideObject::Connector(obj));
        self
    }

    // ─── Group ──────────────────────────────────────────────

    /// Add a group of shapes.
    pub fn add_group(&mut self, children: Vec<SlideObject>, opts: crate::objects::group::GroupOptions) -> &mut Self {
        use crate::enums::EMU;
        let name = self.next_obj_name("Group");
        let cx = opts.position.w.as_ref().map(|c| c.to_emu(12_192_000)).unwrap_or(EMU as i64);
        let cy = opts.position.h.as_ref().map(|c| c.to_emu(6_858_000)).unwrap_or(EMU as i64);
        let child_extent = opts.child_extent.unwrap_or((cx, cy));
        let obj = crate::objects::GroupObject {
            object_name: name,
            position: opts.position,
            child_offset: opts.child_offset,
            child_extent,
            children,
        };
        self.objects.push(SlideObject::Group(obj));
        self
    }

    // ─── Table ──────────────────────────────────────────────

    /// Add a table.
    pub fn add_table(&mut self, rows: Vec<crate::objects::table::TableRow>, opts: crate::objects::table::TableOptions) -> &mut Self {
        let name = self.next_obj_name("Table");
        let mut opts = opts;
        if opts.object_name.is_none() {
            opts.object_name = Some(name.clone());
        }
        let mut rows = rows;
        for row in &mut rows {
            for cell in row.iter_mut() {
                if let Some(ref mut hl) = cell.options.hyperlink {
                    if hl.action.is_none() {
                        self.register_hyperlink_rel(hl);
                    }
                }
            }
        }
        let obj = crate::objects::TableObject { object_name: name, rows, options: opts };
        self.objects.push(SlideObject::Table(obj));
        self
    }
}

/// Optional presentation theme (font faces)
#[derive(Debug, Clone, Default)]
pub struct ThemeProps {
    pub head_font_face: Option<String>,
    pub body_font_face: Option<String>,
}

/// A named section grouping slides
#[derive(Debug, Clone)]
pub struct SectionDef {
    /// Visible section name
    pub name: String,
    /// 1-based slide number of the first slide in this section
    pub start_slide: u32,
}

/// Top-level presentation builder
#[derive(Debug)]
pub struct Presentation {
    pub author: String,
    pub company: String,
    pub title: String,
    pub subject: String,
    pub revision: String,
    pub rtl_mode: bool,
    pub layout: PresLayout,
    pub theme: Option<ThemeProps>,
    pub(crate) slides: Vec<Slide>,
    pub(crate) slide_layouts: Vec<SlideLayout>,
    pub(crate) master: Option<SlideMasterDef>,
    /// Named sections (groupings visible in the slide panel)
    pub(crate) sections: Vec<SectionDef>,
    slide_id_counter: u32,
}

impl Default for Presentation {
    fn default() -> Self {
        Presentation {
            author: String::new(),
            company: String::new(),
            title: String::new(),
            subject: String::new(),
            revision: "1".to_string(),
            rtl_mode: false,
            layout: PresLayout::default(),
            theme: None,
            slides: Vec::new(),
            slide_layouts: vec![SlideLayout { name: "DEFAULT".to_string(), ..Default::default() }],
            master: None,
            sections: Vec::new(),
            slide_id_counter: 0,
        }
    }
}

impl Presentation {
    /// Create a new blank presentation with 16:9 layout
    pub fn new() -> Self {
        Presentation::default()
    }

    /// Create a presentation with a custom layout
    pub fn with_layout(layout: PresLayout) -> Self {
        Presentation { layout, ..Default::default() }
    }

    /// Set the author metadata
    pub fn author(mut self, a: impl Into<String>) -> Self { self.author = a.into(); self }
    /// Set the company metadata
    pub fn company(mut self, c: impl Into<String>) -> Self { self.company = c.into(); self }
    /// Set the title metadata
    pub fn title(mut self, t: impl Into<String>) -> Self { self.title = t.into(); self }
    /// Set the subject metadata
    pub fn subject(mut self, s: impl Into<String>) -> Self { self.subject = s.into(); self }
    /// Set the revision metadata
    pub fn revision(mut self, r: impl Into<String>) -> Self { self.revision = r.into(); self }
    /// Enable right-to-left mode
    pub fn rtl(mut self) -> Self { self.rtl_mode = true; self }
    /// Set a custom theme
    pub fn theme(mut self, t: ThemeProps) -> Self { self.theme = Some(t); self }

    // ─── Layout / Master management ─────────────────────────

    /// Define a custom slide layout with optional dimensions in inches, and
    /// return a mutable reference to it — the same "hand back the thing you
    /// just created, ready to build on" ergonomics [`add_slide`](Presentation::add_slide)
    /// already uses, rather than `&mut Self` chaining (there's nothing useful
    /// to chain: the very next thing every caller does is add chrome objects
    /// to the new layout via [`SlideLayout::add_text`]/[`add_shape`](SlideLayout::add_shape)/etc).
    ///
    /// ```rust,no_run
    /// use deckmint::Presentation;
    /// use deckmint::objects::shape::ShapeOptionsBuilder;
    /// use deckmint::ShapeType;
    ///
    /// let mut pres = Presentation::new();
    /// let layout = pres.define_layout("Fluency Practise", None, None);
    /// layout.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
    ///     .pos(0.0, 0.0).size(3.0, 0.35).fill_color("EEC3E3").build());
    /// ```
    pub fn define_layout(&mut self, name: impl Into<String>, width: Option<f64>, height: Option<f64>) -> &mut SlideLayout {
        self.slide_layouts.push(SlideLayout { name: name.into(), width, height, ..Default::default() });
        self.slide_layouts.last_mut().unwrap()
    }

    /// Set the slide master definition. Objects and background defined here
    /// are rendered into `ppt/slideMasters/slideMaster1.xml`.
    pub fn define_master(&mut self, def: SlideMasterDef) -> &mut Self {
        self.master = Some(def);
        self
    }

    /// Remove a slide by 0-based index and convert it into a [`SlideMasterDef`].
    ///
    /// This lets you build a master using the familiar `Slide` API, then promote it:
    /// ```rust,no_run
    /// use deckmint::{Presentation, ShapeType};
    /// use deckmint::objects::shape::ShapeOptionsBuilder;
    ///
    /// let mut pres = Presentation::new();
    /// let s = pres.add_slide();
    /// s.set_background_color("1F3864");
    /// s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
    ///     .pos(0.0, 7.0).size(10.0, 0.5).fill_color("2E4057").build());
    ///
    /// let master = pres.promote_slide_to_master(0, "Corporate");
    /// pres.define_master(master);
    /// ```
    ///
    /// # Panics
    /// Panics if `idx >= slide_count()`.
    pub fn promote_slide_to_master(&mut self, idx: usize, title: impl Into<String>) -> SlideMasterDef {
        let slide = self.slides.remove(idx);
        slide.into_master(title)
    }

    // ─── Slide management ───────────────────────────────────

    /// Add a named section starting at the next slide to be added.
    /// Sections are visible in the slide panel of PowerPoint / LibreOffice.
    pub fn add_section(&mut self, name: impl Into<String>) -> &mut Self {
        let next_slide_num = self.slide_id_counter + 1;
        self.sections.push(SectionDef { name: name.into(), start_slide: next_slide_num });
        self
    }

    /// Add a new blank slide and return a mutable reference to it.
    ///
    /// The slide binds to `slide_layouts[0]` (the implicit "DEFAULT" layout,
    /// or whichever layout was defined/promoted first) — use
    /// [`add_slide_with_layout`](Presentation::add_slide_with_layout) to bind
    /// to a specific named layout instead.
    pub fn add_slide(&mut self) -> &mut Slide {
        let slide = self.new_slide();
        self.slides.push(slide);
        self.slides.last_mut().unwrap()
    }

    /// Add a new blank slide bound to the named [`SlideLayout`] (defined via
    /// [`define_layout`](Presentation::define_layout)) — the layout's own
    /// chrome objects render underneath whatever this slide adds, and
    /// `ppt/slides/_rels/slideN.xml.rels`'s `rId1` (`slideLayout` relationship)
    /// points at that layout's XML part instead of always `slideLayout1.xml`.
    ///
    /// # Panics
    /// Panics if no layout named `layout_name` has been defined yet (via
    /// [`define_layout`](Presentation::define_layout)) — a slide bound to a
    /// layout that doesn't exist would silently fall back to the wrong chrome
    /// at render time, which is worse than failing loudly at build time.
    pub fn add_slide_with_layout(&mut self, layout_name: &str) -> &mut Slide {
        let idx = self.slide_layouts.iter().position(|l| l.name == layout_name).unwrap_or_else(|| {
            let known: Vec<&str> = self.slide_layouts.iter().map(|l| l.name.as_str()).collect();
            panic!("add_slide_with_layout: no layout named {layout_name:?} — defined layouts: {known:?}")
        });
        let mut slide = self.new_slide();
        slide.layout_idx = Some(idx);
        self.slides.push(slide);
        self.slides.last_mut().unwrap()
    }

    fn new_slide(&mut self) -> Slide {
        self.slide_id_counter += 1;
        let slide_num = self.slide_id_counter;
        let slide_id = 255 + slide_num; // OOXML spec: first slide ID is 256
        let r_id = slide_num + 1;       // rId1 = slide master, rId2 = slide 1, etc.
        Slide::new(slide_num, slide_id, r_id)
    }

    /// Return the number of slides
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Get a reference to a slide by 0-based index
    pub fn slide(&self, idx: usize) -> Option<&Slide> {
        self.slides.get(idx)
    }

    /// Get a mutable reference to a slide by 0-based index
    pub fn slide_mut(&mut self, idx: usize) -> Option<&mut Slide> {
        self.slides.get_mut(idx)
    }

    /// The presentation's defined layouts, in package order (`slide_layouts[0]`
    /// is always the implicit "DEFAULT" layout every fresh `Presentation`
    /// starts with -> `ppt/slideLayouts/slideLayout1.xml`; each
    /// [`define_layout`](Presentation::define_layout) call appends one more).
    /// Read-only mirror of [`slide`](Presentation::slide)/[`slide_count`](Presentation::slide_count)'s
    /// access pattern, for callers that build a whole layout set up front (e.g.
    /// one named layout per pedagogical archetype) and then want to inspect
    /// what got defined.
    pub fn slide_layouts(&self) -> &[SlideLayout] {
        &self.slide_layouts
    }

    // ─── Export ─────────────────────────────────────────────

    /// Assemble the presentation into a ZIP buffer (`Vec<u8>`).
    /// This is synchronous and works on both native and WASM.
    pub fn write(&self) -> Result<Vec<u8>, PptxError> {
        assemble_pptx(self)
    }

    /// Write the presentation to a file (native only).
    #[cfg(not(target_arch = "wasm32"))]
    pub fn write_to_file(&self, path: &str) -> Result<(), PptxError> {
        let bytes = self.write()?;
        std::fs::write(path, &bytes).map_err(PptxError::Io)
    }

    /// Assemble and validate the presentation.
    ///
    /// Returns the PPTX bytes if validation passes, or a
    /// [`PptxError::ValidationFailed`] listing all structural issues.
    pub fn write_validated(&self) -> Result<Vec<u8>, PptxError> {
        let bytes = self.write()?;
        let issues = crate::validate::validate(&bytes);
        let errors: Vec<_> = issues.into_iter()
            .filter(|i| i.severity == crate::validate::Severity::Error)
            .collect();
        if !errors.is_empty() {
            return Err(PptxError::ValidationFailed(errors));
        }
        Ok(bytes)
    }
}
