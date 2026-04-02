use crate::types::{AnimationEffect, Coord, ShapeLineProps};

/// Connector shape type — maps to OOXML prstGeom prst value.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum ConnectorType {
    /// Straight line connector.
    #[default]
    Straight,
    /// Elbow (right-angle) connector.
    Elbow,
    /// Curved connector.
    Curved,
}

impl ConnectorType {
    /// Return the OOXML preset geometry name for this connector type.
    pub fn as_prst(&self) -> &'static str {
        match self {
            ConnectorType::Straight => "line",
            ConnectorType::Elbow    => "bentConnector3",
            ConnectorType::Curved   => "curvedConnector3",
        }
    }
}

/// Optional connection to a specific shape's connection site index.
#[derive(Debug, Clone)]
pub struct ConnectorEndpoint {
    /// The `cNvPr id` of the target shape (assigned as `obj_idx + 2` in the XML loop).
    pub shape_id: u32,
    /// Connection site index on the target shape (0=top, 1=right, 2=bottom, 3=left for most shapes).
    pub site_idx: u32,
}

/// Options controlling connector placement and styling.
#[derive(Debug, Clone, Default)]
pub struct ConnectorOptions {
    /// X coordinate of the start point.
    pub x1: Option<Coord>,
    /// Y coordinate of the start point.
    pub y1: Option<Coord>,
    /// X coordinate of the end point.
    pub x2: Option<Coord>,
    /// Y coordinate of the end point.
    pub y2: Option<Coord>,
    /// Line styling (color, width, dash, arrows).
    pub line: Option<ShapeLineProps>,
    /// Optional connection of the start point to a specific shape.
    pub start_conn: Option<ConnectorEndpoint>,
    /// Optional connection of the end point to a specific shape.
    pub end_conn: Option<ConnectorEndpoint>,
    /// Click-triggered animations on this connector.
    pub animations: Vec<AnimationEffect>,
}

/// A connector line placed on a slide.
#[derive(Debug, Clone)]
pub struct ConnectorObject {
    /// Internal object name for XML identification.
    pub object_name: String,
    /// The type of connector geometry.
    pub connector_type: ConnectorType,
    /// Placement and styling options.
    pub options: ConnectorOptions,
}

/// Builder for ConnectorOptions.
pub struct ConnectorOptionsBuilder {
    /// Options being constructed.
    opts: ConnectorOptions,
    /// Connector geometry type.
    ctype: ConnectorType,
}

impl ConnectorOptionsBuilder {
    /// Create a new builder with default connector options.
    pub fn new() -> Self {
        ConnectorOptionsBuilder { opts: ConnectorOptions::default(), ctype: ConnectorType::Straight }
    }

    /// Set the connector geometry type.
    pub fn connector_type(mut self, t: ConnectorType) -> Self { self.ctype = t; self }
    /// Set the start-point X coordinate in inches.
    pub fn x1(mut self, v: f64) -> Self { self.opts.x1 = Some(Coord::Inches(v)); self }
    /// Set the start-point Y coordinate in inches.
    pub fn y1(mut self, v: f64) -> Self { self.opts.y1 = Some(Coord::Inches(v)); self }
    /// Set the end-point X coordinate in inches.
    pub fn x2(mut self, v: f64) -> Self { self.opts.x2 = Some(Coord::Inches(v)); self }
    /// Set the end-point Y coordinate in inches.
    pub fn y2(mut self, v: f64) -> Self { self.opts.y2 = Some(Coord::Inches(v)); self }
    /// Set the line styling properties.
    pub fn line(mut self, l: ShapeLineProps) -> Self { self.opts.line = Some(l); self }
    /// Connect the start point to a specific shape.
    pub fn start_conn(mut self, e: ConnectorEndpoint) -> Self { self.opts.start_conn = Some(e); self }
    /// Connect the end point to a specific shape.
    pub fn end_conn(mut self, e: ConnectorEndpoint) -> Self { self.opts.end_conn = Some(e); self }
    /// Add a click-triggered animation effect.
    pub fn animation(mut self, a: AnimationEffect) -> Self { self.opts.animations.push(a); self }

    /// Consume the builder and return the connector type and options.
    pub fn build(self) -> (ConnectorType, ConnectorOptions) {
        (self.ctype, self.opts)
    }
}

impl Default for ConnectorOptionsBuilder {
    fn default() -> Self { Self::new() }
}
