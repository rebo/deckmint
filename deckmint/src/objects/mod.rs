//! Slide object types — text boxes, shapes, images, tables, and charts.
//!
//! Each object has an options struct and a corresponding builder for
//! ergonomic construction. Objects are added to a [`crate::Slide`] via
//! methods such as [`crate::Slide::add_text`] and [`crate::Slide::add_chart`].

pub mod chart;
pub mod text;
pub mod shape;
pub mod image;
pub mod table;
pub mod connector;
pub mod media;
pub mod group;

pub use chart::ChartObject;
pub use text::TextObject;
pub use shape::ShapeObject;
pub use image::ImageObject;
pub use table::TableObject;
pub use connector::{ConnectorEndpoint, ConnectorObject, ConnectorOptions, ConnectorOptionsBuilder, ConnectorType};
pub use media::{MediaObject, MediaType};
pub use group::GroupObject;

/// Discriminated union of all slide object types
#[derive(Debug, Clone)]
pub enum SlideObject {
    Text(TextObject),
    Shape(ShapeObject),
    Image(ImageObject),
    Table(TableObject),
    Connector(ConnectorObject),
    Media(MediaObject),
    Group(GroupObject),
}

/// A relationship entry on a slide (hyperlinks, notes)
#[derive(Debug, Clone)]
pub struct SlideRel {
    pub r_id: u32,
    pub rel_type: String,
    pub target: String,
    /// "slide" for internal slide links, otherwise URL
    pub data: Option<String>,
}

/// A media relationship on a slide (images, audio, video)
#[derive(Debug, Clone)]
pub struct SlideRelMedia {
    pub r_id: u32,
    pub rel_type: String,
    pub target: String,
    pub extn: String,
    pub data: Vec<u8>,
}
