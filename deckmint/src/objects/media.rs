use crate::types::{Coord, PositionProps};

/// Type of embedded media.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum MediaType {
    /// Video file (mp4, mov, avi, wmv, etc.)
    Video,
    /// Audio file (mp3, wav, m4a, etc.)
    Audio,
}

/// A media object (video or audio) placed on a slide.
#[derive(Debug, Clone)]
pub struct MediaObject {
    /// Internal object name for relationship tracking.
    pub object_name: String,
    /// Type of media (video or audio).
    pub media_type: MediaType,
    /// rId for the media file relationship.
    pub media_rid: u32,
    /// rId for the poster frame image (video only).
    pub poster_rid: Option<u32>,
    /// File extension of the media (e.g. "mp4", "mp3").
    pub extension: String,
    /// Placement and styling options.
    pub options: MediaOptions,
}

/// Options for media placement and configuration.
#[derive(Debug, Clone)]
pub struct MediaOptions {
    /// Position and dimensions on the slide.
    pub position: PositionProps,
    /// Accessibility description for the media.
    pub alt_text: Option<String>,
}

impl Default for MediaOptions {
    fn default() -> Self {
        MediaOptions {
            position: PositionProps::default(),
            alt_text: None,
        }
    }
}

/// Fluent builder for constructing media placement options.
pub struct MediaOptionsBuilder {
    opts: MediaOptions,
}

impl MediaOptionsBuilder {
    /// Create a new builder with default media options.
    pub fn new() -> Self {
        MediaOptionsBuilder { opts: MediaOptions::default() }
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
    pub fn pos(self, x: f64, y: f64) -> Self { self.x(x).y(y) }
    /// Set size (width, height) in inches.
    pub fn size(self, w: f64, h: f64) -> Self { self.w(w).h(h) }
    /// Set position and size from a [`CellRect`](crate::layout::CellRect).
    pub fn rect(self, r: crate::layout::CellRect) -> Self {
        self.pos(r.x, r.y).size(r.w, r.h)
    }
    /// Set position (x, y) and size (w, h) in inches in a single call.
    pub fn bounds(self, x: f64, y: f64, w: f64, h: f64) -> Self {
        self.pos(x, y).size(w, h)
    }
    /// Set the accessibility alt text.
    pub fn alt_text(mut self, t: impl Into<String>) -> Self { self.opts.alt_text = Some(t.into()); self }

    /// Consume the builder and return the media options.
    pub fn build(self) -> MediaOptions {
        self.opts
    }
}

impl Default for MediaOptionsBuilder {
    fn default() -> Self {
        Self::new()
    }
}
