use crate::types::{AnimationEffect, Coord, HyperlinkProps, ImageColorAdjust, PositionProps, ShadowProps};

/// An image placed on a slide
#[derive(Debug, Clone)]
pub struct ImageObject {
    /// Internal object name for relationship tracking.
    pub object_name: String,
    /// rId in the slide's media relationships.
    pub image_rid: u32,
    /// File extension (png, jpg, gif, svg, etc.).
    pub extension: String,
    /// Actual image bytes (base64-decoded).
    pub data: Vec<u8>,
    /// Whether this image is an SVG (needs special XML handling).
    pub is_svg: bool,
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

    /// Consume the builder and return the options, raw bytes, base64 data, and extension.
    pub fn build(self) -> (ImageOptions, Vec<u8>, Option<String>, Option<String>) {
        (self.opts, self.data, self.data_b64, self.extension)
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
