use crate::objects::{SlideObject, SlideRel, SlideRelMedia, TextObject, ShapeObject, ImageObject, TableObject, ConnectorObject, ConnectorType, ConnectorOptions};
use crate::objects::chart::{ChartObject, ChartOptions, ChartSeries};
use crate::objects::text::{TextRun, TextRunOptions};
use crate::objects::shape::ShapeOptions;
use crate::objects::image::{ImageOptions, ImageOptionsBuilder};
use crate::objects::table::TableOptions;
use crate::enums::{ChartType, ShapeType};
use crate::types::{SlideNumberProps, TransitionProps};

/// A single slide in the presentation
#[derive(Debug, Clone)]
pub struct Slide {
    /// 1-based slide number
    pub slide_num: u32,
    /// Stable ID (starts at 256 per OOXML spec)
    pub slide_id: u32,
    /// rId in presentation.xml.rels (rId2 = slide 1, rId3 = slide 2, etc.)
    pub r_id: u32,
    /// Whether this slide is hidden during presentation.
    pub hidden: bool,
    /// Optional display name for the slide.
    pub name: Option<String>,
    /// Objects placed on this slide (text, shapes, images, tables, connectors).
    pub objects: Vec<SlideObject>,
    /// Non-media relationships (hyperlinks, etc.).
    pub rels: Vec<SlideRel>,
    /// Media relationships (images, audio, video).
    pub rels_media: Vec<SlideRelMedia>,
    /// Charts placed on this slide (separate from spTree objects)
    pub charts: Vec<ChartObject>,
    /// Optional background color or image for this slide.
    pub background: Option<SlideBackground>,
    /// Speaker notes text for this slide.
    pub notes: Option<String>,
    /// Optional slide number placeholder
    pub slide_number: Option<SlideNumberProps>,
    /// Optional slide transition effect.
    pub transition: Option<TransitionProps>,
    /// Which [`SlideLayout`](crate::presentation::SlideLayout) (0-based index
    /// into `Presentation::slide_layouts`) this slide is bound to. `None`
    /// means the first-defined layout (`slide_layouts[0]`, "DEFAULT" unless
    /// overridden) — set via [`Presentation::add_slide_with_layout`], never
    /// directly, so it's always validated against a real defined layout.
    pub(crate) layout_idx: Option<usize>,
    /// Next object counter for auto-naming
    obj_counter: u32,
}

/// Background color or image for a slide
#[derive(Debug, Clone, Default)]
pub struct SlideBackground {
    /// Background hex color, no `#` prefix.
    pub color: Option<String>,
    /// Background color transparency 0–100 (0 = opaque, 100 = fully transparent)
    pub transparency: Option<f64>,
    /// Relationship ID of the background image, if any.
    pub image_rid: Option<u32>,
}

impl Slide {
    pub(crate) fn new(slide_num: u32, slide_id: u32, r_id: u32) -> Self {
        Slide {
            slide_num,
            slide_id,
            r_id,
            hidden: false,
            name: None,
            objects: Vec::new(),
            rels: Vec::new(),
            rels_media: Vec::new(),
            charts: Vec::new(),
            background: None,
            notes: None,
            slide_number: None,
            transition: None,
            layout_idx: None,
            obj_counter: 0,
        }
    }

    fn next_obj_name(&mut self, prefix: &str) -> String {
        self.obj_counter += 1;
        format!("{} {}", prefix, self.obj_counter)
    }

    // ─── add_text ───────────────────────────────────────────

    /// Add a simple text box with the given string and options.
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
        // Register hyperlinks from each run
        let mut runs = runs;
        for run in &mut runs {
            if let Some(ref mut hl) = run.options.hyperlink {
                let rid = self.allocate_rel_rid();
                hl.r_id = rid;
                let (rel_type, target, data) = if let Some(slide) = hl.slide {
                    ("hyperlink".to_string(), slide.to_string(), Some("slide".to_string()))
                } else if let Some(ref url) = hl.url {
                    ("hyperlink".to_string(), url.clone(), None)
                } else {
                    continue;
                };
                self.rels.push(SlideRel { r_id: rid, rel_type, target, data });
            }
        }
        let obj = TextObject { object_name: name, text: runs, options: opts };
        self.objects.push(SlideObject::Text(obj));
        self
    }

    // ─── add_equation ───────────────────────────────────────

    /// Add a text box containing a LaTeX equation.
    ///
    /// The equation is converted to native, editable OMML in PowerPoint.
    /// Requires the `math` feature (enabled by default).
    ///
    /// ```rust,no_run
    /// use deckmint::Presentation;
    /// use deckmint::objects::text::TextOptionsBuilder;
    ///
    /// let mut pres = Presentation::new();
    /// let slide = pres.add_slide();
    /// slide.add_equation(
    ///     r"E = mc^2",
    ///     TextOptionsBuilder::new().pos(1.0, 2.0).size(8.0, 1.0).build(),
    /// ).unwrap();
    /// ```
    #[cfg(feature = "math")]
    pub fn add_equation(
        &mut self,
        latex: &str,
        opts: crate::objects::text::TextOptions,
    ) -> Result<&mut Self, crate::error::PptxError> {
        let run = TextRun::equation(latex)?;
        let name = self.next_obj_name("TextBox");
        let obj = TextObject { object_name: name, text: vec![run], options: opts };
        self.objects.push(SlideObject::Text(obj));
        Ok(self)
    }

    /// Add a text box with mixed plain text and inline LaTeX equations.
    ///
    /// Equations are delimited by `$...$`. Use `\$` to insert a literal dollar sign.
    ///
    /// Plain text segments inherit the font size, color, and face from the
    /// `TextOptions` so the text matches the surrounding formatting.
    ///
    /// ```rust,no_run
    /// use deckmint::Presentation;
    /// use deckmint::objects::text::TextOptionsBuilder;
    ///
    /// let mut pres = Presentation::new();
    /// let slide = pres.add_slide();
    /// slide.add_text_with_math(
    ///     r"If $a = 5$, then $a^2 = 25$.",
    ///     TextOptionsBuilder::new()
    ///         .pos(0.5, 1.5).size(9.0, 1.0)
    ///         .font_size(18.0)
    ///         .build(),
    /// ).unwrap();
    /// ```
    #[cfg(feature = "math")]
    pub fn add_text_with_math(
        &mut self,
        text: &str,
        opts: crate::objects::text::TextOptions,
    ) -> Result<&mut Self, crate::error::PptxError> {
        // Build base run options that inherit from the paragraph-level TextOptions
        let base_run_opts = TextRunOptions {
            font_size: opts.font_size,
            font_face: opts.font_face.clone(),
            color: opts.color.clone(),
            bold: opts.bold,
            italic: opts.italic,
            ..Default::default()
        };

        // Handle \$ escape: replace with a placeholder, split on bare $, then restore
        let escaped = text.replace(r"\$", "\x00");
        let mut runs = Vec::new();
        let mut rest = escaped.as_str();
        while let Some(start) = rest.find('$') {
            // Plain text before the $
            if start > 0 {
                let plain = rest[..start].replace('\x00', "$");
                let mut run = TextRun::new(plain);
                run.options = base_run_opts.clone();
                runs.push(run);
            }
            rest = &rest[start + 1..];
            // Find closing $
            let end = rest.find('$').ok_or_else(|| {
                crate::error::PptxError::InvalidArgument("unmatched $ delimiter in math text".into())
            })?;
            let latex = &rest[..end];
            runs.push(TextRun::equation(latex)?);
            rest = &rest[end + 1..];
        }
        // Remaining plain text
        if !rest.is_empty() {
            let plain = rest.replace('\x00', "$");
            let mut run = TextRun::new(plain);
            run.options = base_run_opts;
            runs.push(run);
        }
        let name = self.next_obj_name("TextBox");
        let obj = TextObject { object_name: name, text: runs, options: opts };
        self.objects.push(SlideObject::Text(obj));
        Ok(self)
    }

    // ─── add_shape ──────────────────────────────────────────

    /// Add a shape.
    pub fn add_shape(&mut self, shape_type: ShapeType, mut opts: ShapeOptions) -> &mut Self {
        // Register click and hover hyperlink rels
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
        let obj = ShapeObject {
            object_name: shape_name,
            shape_type,
            options: opts,
            text: Some(text_obj),
        };
        self.objects.push(SlideObject::Shape(obj));
        self
    }

    // ─── add_image ──────────────────────────────────────────

    /// Add an image from raw bytes.
    pub fn add_image(&mut self, data: Vec<u8>, extension: &str, mut opts: ImageOptions) -> &mut Self {
        let name = self.next_obj_name("Image");

        // Auto-detect image dimensions from bytes if w or h not set, deriving
        // whichever side is missing from the source's intrinsic aspect ratio
        // when the other was given explicitly.
        crate::objects::image::resolve_image_size(&data, &mut opts.position);

        // Register click and hover hyperlink rels
        if let Some(ref mut hl) = opts.hyperlink {
            self.register_hyperlink_rel(hl);
        }
        if let Some(ref mut hl) = opts.hover {
            self.register_hyperlink_rel(hl);
        }

        let rid = self.allocate_media_rid();
        let extn = extension.to_lowercase();
        let is_svg = extn == "svg";
        let target = format!("../media/image{}.{}", rid, extn);

        self.rels_media.push(SlideRelMedia {
            r_id: rid,
            rel_type: "image".to_string(),
            target,
            extn: extn.clone(),
            data: data.clone(),
        });

        let obj = ImageObject { object_name: name, image_rid: rid, extension: extn, data, is_svg, raster_fallback_rid: None, options: opts };
        self.objects.push(SlideObject::Image(obj));
        self
    }

    /// Add an SVG image with a raster (PNG) fallback.
    ///
    /// PowerPoint renders SVG natively via an `asvg:svgBlip` extension on the
    /// picture, but that extension is additive on top of a base
    /// `<a:blip r:embed>` that must point at a real raster image — hosts that
    /// predate SVG support (older PowerPoint, other OOXML consumers) fall
    /// back to it. Both `svg_data` and `png_fallback_data` are registered as
    /// their own independent media relationships (two rIds, two files in
    /// `ppt/media/`); nothing here assumes one is "adjacent" to the other.
    ///
    /// Width/height auto-detection (when `opts` doesn't set both) reads the
    /// **raster fallback's** pixel size, not the SVG's — PNG dimensions are a
    /// plain pixel grid, so they're a more reliable source of truth than
    /// parsing an arbitrary SVG's `viewBox`/`width`/`height` attributes.
    ///
    /// ```rust,no_run
    /// use deckmint::Presentation;
    /// use deckmint::objects::image::ImageOptionsBuilder;
    ///
    /// let mut pres = Presentation::new();
    /// let slide = pres.add_slide();
    /// slide.add_image_svg(
    ///     b"<svg xmlns='http://www.w3.org/2000/svg' .../>".to_vec(),
    ///     vec![/* PNG fallback bytes */],
    ///     ImageOptionsBuilder::new().pos(1.0, 1.0).size(4.0, 3.0).build_opts(),
    /// );
    /// ```
    pub fn add_image_svg(&mut self, svg_data: Vec<u8>, png_fallback_data: Vec<u8>, mut opts: ImageOptions) -> &mut Self {
        let name = self.next_obj_name("Image");

        crate::objects::image::resolve_image_size(&png_fallback_data, &mut opts.position);

        if let Some(ref mut hl) = opts.hyperlink {
            self.register_hyperlink_rel(hl);
        }
        if let Some(ref mut hl) = opts.hover {
            self.register_hyperlink_rel(hl);
        }

        // Raster fallback first — its rId becomes the base <a:blip r:embed> —
        // then the SVG itself, whose rId feeds the <asvg:svgBlip> extension.
        let raster_rid = self.allocate_media_rid();
        self.rels_media.push(SlideRelMedia {
            r_id: raster_rid,
            rel_type: "image".to_string(),
            target: format!("../media/image{raster_rid}.png"),
            extn: "png".to_string(),
            data: png_fallback_data,
        });

        let svg_rid = self.allocate_media_rid();
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
    /// strings (with or without a `data:` URI prefix). See
    /// [`Slide::add_image_svg`].
    pub fn add_image_svg_base64(
        &mut self,
        svg_b64: &str,
        png_fallback_b64: &str,
        opts: ImageOptions,
    ) -> Result<&mut Self, crate::error::PptxError> {
        let svg_data = crate::utils::decode_base64_image(svg_b64)?;
        let fallback_data = crate::utils::decode_base64_image(png_fallback_b64)?;
        Ok(self.add_image_svg(svg_data, fallback_data, opts))
    }

    /// Add an SVG image built via [`ImageOptionsBuilder`] — the builder's
    /// primary data (`.bytes(svg_bytes, "svg")` or `.base64(svg_b64, "svg")`)
    /// supplies the SVG, and `.svg_fallback(png_bytes)` /
    /// `.svg_fallback_base64(png_b64)` supplies the raster fallback. Mirrors
    /// [`Slide::add_image_from`] but for [`Slide::add_image_svg`].
    pub fn add_image_svg_from(&mut self, builder: ImageOptionsBuilder) -> Result<&mut Self, crate::error::PptxError> {
        let (opts, svg_data, svg_b64, fallback_data, fallback_b64) = builder.build_svg();
        let svg_bytes = match svg_b64 {
            Some(b64) => crate::utils::decode_base64_image(&b64)?,
            None => svg_data,
        };
        if svg_bytes.is_empty() {
            return Err(crate::error::PptxError::InvalidArgument(
                "add_image_svg_from: no SVG data; call .bytes(svg_bytes, \"svg\") or .base64(svg_b64, \"svg\") first".to_string(),
            ));
        }
        let fallback_bytes = match fallback_b64 {
            Some(b64) => crate::utils::decode_base64_image(&b64)?,
            None => fallback_data,
        };
        if fallback_bytes.is_empty() {
            return Err(crate::error::PptxError::InvalidArgument(
                "add_image_svg_from: no raster fallback; call .svg_fallback(png_bytes) or .svg_fallback_base64(...) — PowerPoint's base <a:blip> requires a real raster image, not just the SVG".to_string(),
            ));
        }
        Ok(self.add_image_svg(svg_bytes, fallback_bytes, opts))
    }

    /// Add an image from a base64-encoded string (with or without data: prefix).
    pub fn add_image_base64(&mut self, b64: &str, extension: &str, opts: ImageOptions) -> Result<&mut Self, crate::error::PptxError> {
        // Strip data URI prefix if present
        let raw = if let Some(idx) = b64.find(',') {
            &b64[idx + 1..]
        } else {
            b64
        };
        let data = base64::Engine::decode(&base64::engine::general_purpose::STANDARD, raw)
            .map_err(|e| crate::error::PptxError::InvalidArgument(format!("base64 decode error: {e}")))?;
        Ok(self.add_image(data, extension, opts))
    }

    /// Add an image using an [`ImageOptionsBuilder`] directly.
    ///
    /// The builder must have image data set via `.bytes()` or `.base64()`.
    ///
    /// ```rust,no_run
    /// use deckmint::Presentation;
    /// use deckmint::objects::image::ImageOptionsBuilder;
    ///
    /// let mut pres = Presentation::new();
    /// let slide = pres.add_slide();
    /// slide.add_image_from(
    ///     ImageOptionsBuilder::new()
    ///         .bytes(vec![/* png bytes */], "png")
    ///         .pos(1.0, 1.0).size(4.0, 3.0)
    /// ).unwrap();
    /// ```
    pub fn add_image_from(&mut self, builder: ImageOptionsBuilder) -> Result<&mut Self, crate::error::PptxError> {
        let (opts, data, data_b64, ext) = builder.build();
        if let Some(b64_str) = data_b64 {
            let extension = ext.unwrap_or_else(|| "png".to_string());
            return self.add_image_base64(&b64_str, &extension, opts);
        }
        if data.is_empty() {
            return Err(crate::error::PptxError::InvalidArgument(
                "ImageOptionsBuilder has no image data; call .bytes() or .base64() first".to_string(),
            ));
        }
        let extension = ext.unwrap_or_else(|| "png".to_string());
        Ok(self.add_image(data, &extension, opts))
    }

    // ─── add_video / add_audio ────────────────────────────

    /// Add a video to this slide.
    ///
    /// `data` is the raw video bytes, `extension` the file format (e.g. `"mp4"`).
    /// `poster_data` / `poster_ext` provide a poster frame image shown before playback.
    pub fn add_video(
        &mut self,
        data: Vec<u8>,
        extension: &str,
        poster_data: Vec<u8>,
        poster_ext: &str,
        opts: crate::objects::media::MediaOptions,
    ) -> &mut Self {
        let name = self.next_obj_name("Video");
        let extn = extension.to_lowercase();
        let poster_extn = poster_ext.to_lowercase();

        // Allocate rId for poster image first
        let poster_rid = self.allocate_media_rid();
        let poster_target = format!("../media/image{}.{}", poster_rid, poster_extn);
        self.rels_media.push(SlideRelMedia {
            r_id: poster_rid,
            rel_type: "image".to_string(),
            target: poster_target,
            extn: poster_extn,
            data: poster_data,
        });

        // Allocate rId for video file
        let media_rid = self.allocate_media_rid();
        let media_target = format!("../media/media{}.{}", media_rid, extn);
        self.rels_media.push(SlideRelMedia {
            r_id: media_rid,
            rel_type: "video".to_string(),
            target: media_target,
            extn,
            data,
        });

        let obj = crate::objects::MediaObject {
            object_name: name,
            media_type: crate::objects::MediaType::Video,
            media_rid,
            poster_rid: Some(poster_rid),
            extension: extension.to_lowercase(),
            options: opts,
        };
        self.objects.push(SlideObject::Media(obj));
        self
    }

    /// Add an audio file to this slide.
    ///
    /// `data` is the raw audio bytes, `extension` the file format (e.g. `"mp3"`).
    pub fn add_audio(
        &mut self,
        data: Vec<u8>,
        extension: &str,
        opts: crate::objects::media::MediaOptions,
    ) -> &mut Self {
        let name = self.next_obj_name("Audio");
        let extn = extension.to_lowercase();

        let media_rid = self.allocate_media_rid();
        let media_target = format!("../media/media{}.{}", media_rid, extn);
        self.rels_media.push(SlideRelMedia {
            r_id: media_rid,
            rel_type: "audio".to_string(),
            target: media_target,
            extn,
            data,
        });

        let obj = crate::objects::MediaObject {
            object_name: name,
            media_type: crate::objects::MediaType::Audio,
            media_rid,
            poster_rid: None,
            extension: extension.to_lowercase(),
            options: opts,
        };
        self.objects.push(SlideObject::Media(obj));
        self
    }

    // ─── add_group ─────────────────────────────────────────

    /// Add a group of shapes that act as a single unit.
    ///
    /// Child shapes should be pre-built as `SlideObject` variants.
    /// The group's child coordinate space defaults to matching the group's own dimensions.
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

    fn allocate_media_rid(&mut self) -> u32 {
        // rId1 = slideLayout, rId2 = notesSlide are reserved; user content starts at rId3
        (self.rels.len() + self.rels_media.len() + self.charts.len() + 3) as u32
    }

    fn allocate_rel_rid(&mut self) -> u32 {
        // rId1 = slideLayout, rId2 = notesSlide are reserved; user content starts at rId3
        (self.rels.len() + self.rels_media.len() + self.charts.len() + 3) as u32
    }

    /// Register a hyperlink's relationship (if it needs one) and set its r_id.
    fn register_hyperlink_rel(&mut self, hl: &mut crate::types::HyperlinkProps) {
        if hl.action.is_none() {
            let rid = self.allocate_rel_rid();
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

    fn allocate_chart_rid(&mut self) -> u32 {
        // rId1 = slideLayout, rId2 = notesSlide are reserved; user content starts at rId3
        (self.rels.len() + self.rels_media.len() + self.charts.len() + 3) as u32
    }

    // ─── add_chart ──────────────────────────────────────────

    /// Add a chart to this slide.
    pub fn add_chart(&mut self, chart_type: ChartType, series: Vec<ChartSeries>, opts: ChartOptions) -> &mut Self {
        let name = self.next_obj_name("Chart");
        let chart_rid = self.allocate_chart_rid();
        let chart = ChartObject { object_name: name, chart_rid, chart_type, series, options: opts };
        self.charts.push(chart);
        self
    }

    // ─── add_table ──────────────────────────────────────────

    /// Add a table.
    pub fn add_table(&mut self, rows: Vec<crate::objects::table::TableRow>, opts: TableOptions) -> &mut Self {
        let name = self.next_obj_name("Table");
        let mut opts = opts;
        if opts.object_name.is_none() {
            opts.object_name = Some(name.clone());
        }
        // Register cell hyperlink rels (navigation actions need no rel)
        let mut rows = rows;
        for row in &mut rows {
            for cell in row.iter_mut() {
                if let Some(ref mut hl) = cell.options.hyperlink {
                    if hl.action.is_none() {
                        let rid = self.allocate_rel_rid();
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
            }
        }
        let obj = TableObject { object_name: name, rows, options: opts };
        self.objects.push(SlideObject::Table(obj));
        self
    }

    // ─── Speaker notes ──────────────────────────────────────

    /// Add speaker notes to this slide.
    pub fn add_notes(&mut self, notes: impl Into<String>) -> &mut Self {
        self.notes = Some(notes.into());
        self
    }

    // ─── Connectors ─────────────────────────────────────────

    /// Add a connector line between two points.
    pub fn add_connector(&mut self, connector_type: ConnectorType, opts: ConnectorOptions) -> &mut Self {
        let name = self.next_obj_name("Connector");
        let obj = ConnectorObject { object_name: name, connector_type, options: opts };
        self.objects.push(SlideObject::Connector(obj));
        self
    }

    // ─── Transition ─────────────────────────────────────────

    /// Set the slide transition effect.
    pub fn set_transition(&mut self, props: TransitionProps) -> &mut Self {
        self.transition = Some(props);
        self
    }

    // ─── Background ─────────────────────────────────────────

    /// Set the background color as 6-digit hex, no `#` prefix.
    pub fn set_background_color(&mut self, color: impl Into<String>) -> &mut Self {
        let bg = self.background.get_or_insert_with(SlideBackground::default);
        bg.color = Some(color.into().trim_start_matches('#').to_uppercase());
        self
    }

    /// Set background color with transparency (0 = opaque, 100 = fully transparent).
    pub fn set_background_color_transparency(&mut self, color: impl Into<String>, transparency: f64) -> &mut Self {
        let bg = self.background.get_or_insert_with(SlideBackground::default);
        bg.color = Some(color.into().trim_start_matches('#').to_uppercase());
        bg.transparency = Some(transparency);
        self
    }

    /// Set a background image for this slide from raw bytes.
    pub fn set_background_image(&mut self, data: Vec<u8>, extension: &str) -> &mut Self {
        let rid = self.allocate_media_rid();
        let extn = extension.to_lowercase();
        let target = format!("../media/image{}.{}", rid, extn);
        self.rels_media.push(SlideRelMedia {
            r_id: rid,
            rel_type: "image".to_string(),
            target,
            extn,
            data,
        });
        let bg = self.background.get_or_insert_with(SlideBackground::default);
        bg.image_rid = Some(rid);
        self
    }

    /// Add a slide number placeholder at the given position.
    /// The number is rendered automatically by PowerPoint/LibreOffice.
    pub fn set_slide_number(&mut self, props: SlideNumberProps) -> &mut Self {
        self.slide_number = Some(props);
        self
    }

    /// Hide this slide during presentation playback.
    pub fn hide(&mut self) -> &mut Self {
        self.hidden = true;
        self
    }

    /// Convert this slide into a [`SlideMasterDef`](crate::presentation::SlideMasterDef),
    /// transferring all objects, media, and background settings.
    pub fn into_master(self, title: impl Into<String>) -> crate::presentation::SlideMasterDef {
        crate::presentation::SlideMasterDef {
            title: title.into(),
            background_color: self.background.as_ref().and_then(|b| b.color.clone()),
            background_transparency: self.background.as_ref().and_then(|b| b.transparency),
            background_image_rid: self.background.as_ref().and_then(|b| b.image_rid),
            objects: self.objects,
            rels: self.rels,
            rels_media: self.rels_media,
            obj_counter: self.obj_counter,
        }
    }
}
