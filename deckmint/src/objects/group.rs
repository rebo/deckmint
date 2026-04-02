use crate::objects::SlideObject;
use crate::types::{Coord, PositionProps};

/// A group of shapes that act as a single unit on a slide.
///
/// Child shapes are positioned relative to the group's child coordinate space
/// defined by `child_offset` and `child_extent`.
#[derive(Debug, Clone)]
pub struct GroupObject {
    /// Internal object name.
    pub object_name: String,
    /// Position and dimensions of the group bounding box.
    pub position: PositionProps,
    /// Child coordinate space origin (x, y) in EMU.
    pub child_offset: (i64, i64),
    /// Child coordinate space extent (cx, cy) in EMU.
    pub child_extent: (i64, i64),
    /// Child shapes within this group.
    pub children: Vec<SlideObject>,
}

/// Options for group shape construction.
#[derive(Debug, Clone)]
pub struct GroupOptions {
    /// Position and dimensions of the group bounding box.
    pub position: PositionProps,
    /// Child coordinate space origin (x, y) in EMU.
    /// Defaults to (0, 0) — children positioned from top-left.
    pub child_offset: (i64, i64),
    /// Child coordinate space extent (cx, cy) in EMU.
    /// Defaults to match group dimensions.
    pub child_extent: Option<(i64, i64)>,
}

impl Default for GroupOptions {
    fn default() -> Self {
        GroupOptions {
            position: PositionProps::default(),
            child_offset: (0, 0),
            child_extent: None,
        }
    }
}

/// Fluent builder for group options.
pub struct GroupOptionsBuilder {
    opts: GroupOptions,
}

impl GroupOptionsBuilder {
    /// Create a new builder with default group options.
    pub fn new() -> Self {
        GroupOptionsBuilder { opts: GroupOptions::default() }
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
    /// Set the child coordinate space origin in EMU.
    pub fn child_offset(mut self, x: i64, y: i64) -> Self {
        self.opts.child_offset = (x, y);
        self
    }
    /// Set the child coordinate space extent in EMU.
    /// If not set, defaults to the group's own dimensions.
    pub fn child_extent(mut self, cx: i64, cy: i64) -> Self {
        self.opts.child_extent = Some((cx, cy));
        self
    }

    /// Consume the builder and return the group options.
    pub fn build(self) -> GroupOptions {
        self.opts
    }
}

impl Default for GroupOptionsBuilder {
    fn default() -> Self {
        Self::new()
    }
}
