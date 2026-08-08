use imagesize;

use crate::types::{AnimationEffect, Coord, HyperlinkProps, ImageColorAdjust, PositionProps, ShadowProps};

/// An image placed on a slide
#[derive(Debug, Clone)]
pub struct ImageObject {
    /// Internal object name for relationship tracking.
    pub object_name: String,
    /// rId in the slide's media relationships.
    ///
    /// For a raster image this is the only media rId, referenced by the
    /// picture's base `<a:blip r:embed>`. For an SVG image (`is_svg == true`,
    /// built via `add_image_svg`) this is instead the **SVG's own** media rId
    /// — referenced from the `<asvg:svgBlip r:embed>` extension — while the
    /// raster fallback PowerPoint's base `<a:blip r:embed>` points at lives in
    /// [`raster_fallback_rid`].
    pub image_rid: u32,
    /// File extension (png, jpg, gif, svg, etc.). For an SVG image this is
    /// always `"svg"` — the raster fallback's own extension (always `"png"`)
    /// isn't tracked here since it isn't rendered from this field.
    pub extension: String,
    /// Actual image bytes (base64-decoded). For an SVG image this is the SVG
    /// source; the raster fallback's bytes live only in the slide/layout/
    /// master's `rels_media` (registered separately by `add_image_svg`).
    pub data: Vec<u8>,
    /// Whether this image is an SVG (needs special XML handling — a raster
    /// `<a:blip>` PLUS an `asvg:svgBlip` extension pointing at the SVG rel).
    pub is_svg: bool,
    /// rId of the raster (PNG) fallback media, present only when `is_svg` is
    /// `true`. PowerPoint's base `<a:blip r:embed>` on an SVG picture must
    /// reference a real, independently-addressable raster image — the
    /// `asvg:svgBlip` extension is additive, not a replacement, so this rId
    /// has to point at its own media part, never reuse or guess at
    /// `image_rid - 1`.
    pub raster_fallback_rid: Option<u32>,
    /// Placement and styling options for this image.
    pub options: ImageOptions,
}

/// Options for image placement
#[derive(Debug, Clone)]
pub struct ImageOptions {
    /// Position and dimensions on the slide.
    pub position: PositionProps,
    /// Accessibility description for the image.
    pub alt_text: Option<String>,
    /// Image sizing and cropping configuration.
    pub sizing: Option<ImageSizing>,
    /// Transparency level, 0.0 (opaque) to 100.0 (fully transparent).
    pub transparency: Option<f64>,
    /// Whether to apply rounded-corner clipping to the image.
    pub rounding: bool,
    /// Drop shadow effect applied to the image.
    pub shadow: Option<ShadowProps>,
    /// Hyperlink activated when the image is clicked.
    pub hyperlink: Option<HyperlinkProps>,
    /// Hyperlink activated when the image is hovered over.
    pub hover: Option<HyperlinkProps>,
    /// Rotation in degrees (clockwise).
    pub rotate: Option<f64>,
    /// Flip horizontally.
    pub flip_h: bool,
    /// Flip vertically.
    pub flip_v: bool,
    /// Colour adjustments (brightness, contrast, greyscale).
    pub color_adjust: Option<ImageColorAdjust>,
    /// Click-triggered animations on this image (each fires on its own click).
    pub animations: Vec<AnimationEffect>,
    /// LTRB crop values as fractions 0.0–1.0 [left, top, right, bottom].
    /// Applied via `<a:srcRect>` when no `sizing` mode is set.
    pub crop: Option<[f64; 4]>,
}

impl Default for ImageOptions {
    fn default() -> Self {
        ImageOptions {
            position: PositionProps::default(),
            alt_text: None,
            sizing: None,
            transparency: None,
            rounding: false,
            shadow: None,
            hyperlink: None,
            hover: None,
            rotate: None,
            flip_h: false,
            flip_v: false,
            color_adjust: None,
            animations: Vec::new(),
            crop: None,
        }
    }
}

/// Image sizing/cropping mode
#[derive(Debug, Clone)]
pub struct ImageSizing {
    /// Sizing strategy (cover, contain, or crop).
    pub sizing_type: ImageSizingType,
    /// Crop offset X in inches.
    pub x: Option<f64>,
    /// Crop offset Y in inches.
    pub y: Option<f64>,
    /// Target width in inches.
    pub w: Option<f64>,
    /// Target height in inches.
    pub h: Option<f64>,
}

/// Auto-fill missing width/height on `position` from the intrinsic pixel size
/// of `data` (96 DPI), preserving the intrinsic aspect ratio when only one of
/// width/height is explicitly set.
///
/// - Neither set: both come straight from the intrinsic pixel size, as before.
/// - Only one set (the common "give me a natural width, keep proportions"
///   case): the other is now *derived* from the intrinsic aspect ratio scaled
///   to match the side that was given — not filled in with the raw intrinsic
///   size regardless of what the caller actually asked for. (That was the bug:
///   `.w(4.0)` alone on a 2:1 image used to still emit the source's raw
///   intrinsic height, silently stretching or squashing the picture instead of
///   scaling it.)
/// - Both set: both are honoured exactly as given — giving a `w`/`h` pair
///   whose ratio doesn't match the source image will visibly distort it;
///   that's on the caller.
/// - A `Coord::Percent` side can't be resolved to EMU here (that needs the
///   slide's own dimensions, unavailable at `add_image` time), so a percent
///   side filling in for a missing partner falls back to the raw intrinsic
///   size for that partner rather than a derived one.
pub(crate) fn resolve_image_size(data: &[u8], position: &mut PositionProps) {
    if position.w.is_some() && position.h.is_some() {
        return;
    }
    let Ok(dim) = imagesize::blob_size(data) else { return };
    if dim.width == 0 || dim.height == 0 {
        return;
    }
    // Assume 96 DPI for pixel→EMU conversion, matching the rest of the crate.
    let emu_w = (dim.width as i64) * 914_400 / 96;
    let emu_h = (dim.height as i64) * 914_400 / 96;

    match (&position.w, &position.h) {
        (None, None) => {
            position.w = Some(Coord::Emu(emu_w));
            position.h = Some(Coord::Emu(emu_h));
        }
        (Some(Coord::Percent(_)), None) => {
            position.h = Some(Coord::Emu(emu_h));
        }
        (None, Some(Coord::Percent(_))) => {
            position.w = Some(Coord::Emu(emu_w));
        }
        (Some(w), None) => {
            let w_emu = w.to_emu(0);
            let derived_h = (w_emu as f64 * emu_h as f64 / emu_w as f64).round() as i64;
            position.h = Some(Coord::Emu(derived_h));
        }
        (None, Some(h)) => {
            let h_emu = h.to_emu(0);
            let derived_w = (h_emu as f64 * emu_w as f64 / emu_h as f64).round() as i64;
            position.w = Some(Coord::Emu(derived_w));
        }
        (Some(_), Some(_)) => unreachable!("handled by the early return above"),
    }
}

/// Strategy for fitting an image within its bounding box.
#[derive(Debug, Clone, PartialEq)]
pub enum ImageSizingType {
    /// Scale to fill the entire area, cropping excess.
    Cover,
    /// Scale to fit within the area, preserving aspect ratio.
    Contain,
    /// Display a cropped region of the original image.
    Crop,
}

/// Fluent builder for constructing image placement options.
pub struct ImageOptionsBuilder {
    opts: ImageOptions,
    /// Raw image bytes.
    pub data: Vec<u8>,
    /// File path (for native feature).
    pub path: Option<String>,
    /// Base64-encoded data string (with or without data: prefix).
    pub data_b64: Option<String>,
    /// File extension override (e.g. "png", "jpg").
    pub extension: Option<String>,
    /// Raster (PNG) fallback bytes for an SVG image, set via
    /// [`svg_fallback`](Self::svg_fallback). Only meaningful alongside
    /// `.bytes(svg_bytes, "svg")` / `.base64(svg_b64, "svg")`, and consumed by
    /// `add_image_svg_from` (not `add_image_from`, which knows nothing about
    /// SVG's two-media-part requirement).
    pub svg_fallback_data: Vec<u8>,
    /// Base64-encoded raster fallback, set via
    /// [`svg_fallback_base64`](Self::svg_fallback_base64).
    pub svg_fallback_data_b64: Option<String>,
}

impl ImageOptionsBuilder {
    /// Create a new builder with default image options.
    pub fn new() -> Self {
        ImageOptionsBuilder {
            opts: ImageOptions::default(),
            data: Vec::new(),
            path: None,
            data_b64: None,
            extension: None,
            svg_fallback_data: Vec::new(),
            svg_fallback_data_b64: None,
        }
    }

    /// Set the X position in inches.
    pub fn x(mut self, v: f64) -> Self { self.opts.position.x = Some(Coord::Inches(v)); self }
    /// Set the Y position in inches.
    pub fn y(mut self, v: f64) -> Self { self.opts.position.y = Some(Coord::Inches(v)); self }
    /// Set the width in inches.
    pub fn w(mut self, v: f64) -> Self { self.opts.position.w = Some(Coord::Inches(v)); self }
    /// Set the height in inches.
    pub fn h(mut self, v: f64) -> Self { self.opts.position.h = Some(Coord::Inches(v)); self }
    /// Set position (x, y) in inches.
    pub fn pos(self, x: f64, y: f64) -> Self {
        self.x(x).y(y)
    }
    /// Set size (width, height) in inches.
    pub fn size(self, w: f64, h: f64) -> Self {
        self.w(w).h(h)
    }
    /// Set position and size from a [`CellRect`](crate::layout::CellRect).
    pub fn rect(self, r: crate::layout::CellRect) -> Self {
        self.pos(r.x, r.y).size(r.w, r.h)
    }
    /// Set position (x, y) and size (w, h) in inches in a single call.
    pub fn bounds(self, x: f64, y: f64, w: f64, h: f64) -> Self {
        self.pos(x, y).size(w, h)
    }
    /// Set the X position as a percentage of slide width.
    pub fn x_pct(mut self, v: f64) -> Self { self.opts.position.x = Some(Coord::Percent(v)); self }
    /// Set the Y position as a percentage of slide height.
    pub fn y_pct(mut self, v: f64) -> Self { self.opts.position.y = Some(Coord::Percent(v)); self }
    /// Set the width as a percentage of slide width.
    pub fn w_pct(mut self, v: f64) -> Self { self.opts.position.w = Some(Coord::Percent(v)); self }
    /// Set the height as a percentage of slide height.
    pub fn h_pct(mut self, v: f64) -> Self { self.opts.position.h = Some(Coord::Percent(v)); self }

    /// Set raw image bytes + extension
    pub fn bytes(mut self, data: Vec<u8>, ext: impl Into<String>) -> Self {
        self.data = data;
        self.extension = Some(ext.into());
        self
    }

    /// Set base64-encoded image data (e.g. "image/png;base64,...")
    pub fn base64(mut self, b64: impl Into<String>, ext: impl Into<String>) -> Self {
        self.data_b64 = Some(b64.into());
        self.extension = Some(ext.into());
        self
    }

    /// Set the accessibility alt text for the image.
    pub fn alt_text(mut self, t: impl Into<String>) -> Self { self.opts.alt_text = Some(t.into()); self }
    /// Set the transparency level, 0.0 (opaque) to 100.0 (fully transparent).
    pub fn transparency(mut self, t: f64) -> Self { self.opts.transparency = Some(t); self }
    /// Enable rounded-corner clipping on the image.
    pub fn rounding(mut self) -> Self { self.opts.rounding = true; self }
    /// Apply a drop shadow effect to the image.
    pub fn shadow(mut self, s: ShadowProps) -> Self { self.opts.shadow = Some(s); self }
    /// Attach a hyperlink activated on click.
    pub fn hyperlink(mut self, h: HyperlinkProps) -> Self { self.opts.hyperlink = Some(h); self }
    /// Attach a hyperlink activated on hover.
    pub fn hover(mut self, h: HyperlinkProps) -> Self { self.opts.hover = Some(h); self }
    /// Set the clockwise rotation in degrees.
    pub fn rotate(mut self, deg: f64) -> Self { self.opts.rotate = Some(deg); self }
    /// Flip the image horizontally.
    pub fn flip_h(mut self) -> Self { self.opts.flip_h = true; self }
    /// Flip the image vertically.
    pub fn flip_v(mut self) -> Self { self.opts.flip_v = true; self }
    /// Add a click-triggered animation effect.
    pub fn animation(mut self, anim: AnimationEffect) -> Self { self.opts.animations.push(anim); self }
    /// Adjust brightness (−100.0 to +100.0, 0.0 = no change).
    pub fn brightness(mut self, v: f64) -> Self {
        self.opts.color_adjust.get_or_insert_with(Default::default).brightness = Some(v); self
    }
    /// Adjust contrast (−100.0 to +100.0, 0.0 = no change).
    pub fn contrast(mut self, v: f64) -> Self {
        self.opts.color_adjust.get_or_insert_with(Default::default).contrast = Some(v); self
    }
    /// Convert image to greyscale.
    pub fn grayscale(mut self) -> Self {
        self.opts.color_adjust.get_or_insert_with(Default::default).grayscale = true; self
    }
    /// Crop the image by specifying fractions (0.0–1.0) to remove from each side.
    /// For example, `crop(0.1, 0.0, 0.1, 0.0)` removes 10% from left and right.
    pub fn crop(mut self, left: f64, top: f64, right: f64, bottom: f64) -> Self {
        self.opts.crop = Some([left, top, right, bottom]);
        self
    }

    /// Set the raster fallback image bytes for an SVG (used by
    /// `add_image_svg_from` and its `SlideLayout`/`SlideMasterDef`
    /// equivalents). PowerPoint's base `<a:blip>` always points at this
    /// raster fallback; the SVG itself is referenced through the
    /// `asvg:svgBlip` extension and takes over rendering wherever the host
    /// supports it.
    pub fn svg_fallback(mut self, data: Vec<u8>) -> Self {
        self.svg_fallback_data = data;
        self
    }

    /// Set the raster fallback as base64-encoded data (with or without a
    /// `data:` URI prefix). See [`svg_fallback`](Self::svg_fallback).
    pub fn svg_fallback_base64(mut self, b64: impl Into<String>) -> Self {
        self.svg_fallback_data_b64 = Some(b64.into());
        self
    }

    /// Consume the builder and return the options, raw bytes, base64 data, and extension.
    pub fn build(self) -> (ImageOptions, Vec<u8>, Option<String>, Option<String>) {
        (self.opts, self.data, self.data_b64, self.extension)
    }

    /// Consume the builder and return the options plus both the primary
    /// (SVG) data and the raster fallback data, each either as raw bytes or
    /// base64 — for `add_image_svg_from` and its `SlideLayout`/
    /// `SlideMasterDef` equivalents.
    pub fn build_svg(self) -> (ImageOptions, Vec<u8>, Option<String>, Vec<u8>, Option<String>) {
        (self.opts, self.data, self.data_b64, self.svg_fallback_data, self.svg_fallback_data_b64)
    }

    /// Consume the builder and return only the image options.
    ///
    /// Use this when you pass image data separately to [`crate::Slide::add_image`].
    pub fn build_opts(self) -> ImageOptions {
        self.opts
    }
}

impl Default for ImageOptionsBuilder {
    fn default() -> Self {
        Self::new()
    }
}
