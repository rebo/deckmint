use crate::enums::ShapeType;
use crate::types::{AnimationEffect, Color, Coord, CustomGeomPoint, GradientFill, HyperlinkProps, LineCap, LineJoin, PatternFill, PositionProps, Scene3DProps, ShadowProps, Shape3DProps, ShapeFillProps, ShapeLineProps};
use crate::objects::text::TextObject;

/// A shape placed on a slide
#[derive(Debug, Clone)]
pub struct ShapeObject {
    /// Internal object name for relationship tracking.
    pub object_name: String,
    /// Preset geometry type for the shape.
    pub shape_type: ShapeType,
    /// Placement and styling options for this shape.
    pub options: ShapeOptions,
    /// Optional text body inside the shape.
    pub text: Option<TextObject>,
}

/// Options for shape placement and styling
#[derive(Debug, Clone)]
pub struct ShapeOptions {
    /// Position and dimensions on the slide.
    pub position: PositionProps,
    /// Fill color or gradient for the shape interior.
    pub fill: Option<ShapeFillProps>,
    /// Outline stroke color and width.
    pub line: Option<ShapeLineProps>,
    /// Drop shadow effect applied to the shape.
    pub shadow: Option<ShadowProps>,
    /// Clockwise rotation in degrees.
    pub rotate: Option<f64>,
    /// Flip the shape horizontally.
    pub flip_h: bool,
    /// Flip the shape vertically.
    pub flip_v: bool,
    /// Corner radius in inches for rounded rectangle shapes.
    pub rect_radius: Option<f64>,
    /// Optional hyperlink on the shape (URL or slide jump).
    pub hyperlink: Option<HyperlinkProps>,
    /// Optional hover action on the shape (URL or slide jump).
    pub hover: Option<HyperlinkProps>,
    /// Alt text / accessibility description (sets `descr` on `<p:cNvPr>`)
    pub alt_text: Option<String>,
    /// [startAngle, swingAngle] in degrees for PIE / ARC / BLOCK_ARC shapes.
    /// Example: [0.0, 270.0] = a three-quarter circle starting from 3 o'clock.
    pub angle_range: Option<[f64; 2]>,
    /// Inner-radius ratio 0.0–1.0 for BLOCK_ARC (default ~0.5 if omitted).
    pub arc_thickness: Option<f64>,
    /// Custom freeform geometry. When set, overrides `shape_type` preset geometry
    /// and emits `<a:custGeom>` instead of `<a:prstGeom>`.
    pub custom_geometry: Option<Vec<CustomGeomPoint>>,
    /// Click-triggered animations on this shape (each fires on its own click).
    pub animations: Vec<AnimationEffect>,
    /// 3D shape properties (bevel, extrusion, material).
    pub shape_3d: Option<Shape3DProps>,
    /// 3D scene properties (camera, light rig).
    pub scene_3d: Option<Scene3DProps>,
}

impl Default for ShapeOptions {
    fn default() -> Self {
        ShapeOptions {
            position: PositionProps::default(),
            fill: None,
            line: None,
            shadow: None,
            rotate: None,
            flip_h: false,
            flip_v: false,
            rect_radius: None,
            hyperlink: None,
            hover: None,
            alt_text: None,
            angle_range: None,
            arc_thickness: None,
            custom_geometry: None,
            animations: Vec::new(),
            shape_3d: None,
            scene_3d: None,
        }
    }
}

/// Builder for shape options
pub struct ShapeOptionsBuilder {
    opts: ShapeOptions,
}

impl ShapeOptionsBuilder {
    /// Create a new builder with default shape options.
    pub fn new() -> Self {
        ShapeOptionsBuilder { opts: ShapeOptions::default() }
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

    /// Set a solid fill color, 6-digit hex, no `#` prefix.
    pub fn fill_color(mut self, color: impl Into<String>) -> Self {
        let mut fill = ShapeFillProps::default();
        fill.color = Some(Color::Hex(color.into().trim_start_matches('#').to_uppercase()));
        self.opts.fill = Some(fill);
        self
    }

    /// Set a solid fill using a `Color` value (supports hex and theme colours).
    pub fn fill_color_value(mut self, color: Color) -> Self {
        let mut fill = ShapeFillProps::default();
        fill.color = Some(color);
        self.opts.fill = Some(fill);
        self
    }

    /// Remove fill entirely, making the shape transparent.
    pub fn no_fill(mut self) -> Self {
        let mut fill = ShapeFillProps::default();
        fill.fill_type = crate::types::FillType::None;
        self.opts.fill = Some(fill);
        self
    }

    /// Apply a gradient fill to the shape.
    pub fn gradient_fill(mut self, g: GradientFill) -> Self {
        self.opts.fill = Some(ShapeFillProps {
            fill_type: crate::types::FillType::Gradient,
            color: None,
            transparency: None,
            gradient: Some(g),
            pattern: None,
        });
        self
    }

    /// Set the outline color, 6-digit hex, no `#` prefix.
    pub fn line_color(mut self, color: impl Into<String>) -> Self {
        let line = self.opts.line.get_or_insert_with(ShapeLineProps::default);
        line.color = Some(color.into().trim_start_matches('#').to_uppercase());
        self
    }

    /// Set the outline width in points.
    pub fn line_width(mut self, pt: f64) -> Self {
        let line = self.opts.line.get_or_insert_with(ShapeLineProps::default);
        line.width = Some(pt);
        self
    }

    /// Set the line cap style (flat, round, or square).
    pub fn line_cap(mut self, cap: LineCap) -> Self {
        let line = self.opts.line.get_or_insert_with(ShapeLineProps::default);
        line.cap = Some(cap);
        self
    }

    /// Set the line join style (round, bevel, or miter).
    pub fn line_join(mut self, join: LineJoin) -> Self {
        let line = self.opts.line.get_or_insert_with(ShapeLineProps::default);
        line.join = Some(join);
        self
    }

    /// Set the clockwise rotation in degrees.
    pub fn rotate(mut self, deg: f64) -> Self { self.opts.rotate = Some(deg); self }
    /// Flip the shape horizontally.
    pub fn flip_h(mut self) -> Self { self.opts.flip_h = true; self }
    /// Flip the shape vertically.
    pub fn flip_v(mut self) -> Self { self.opts.flip_v = true; self }
    /// Apply a drop shadow effect to the shape.
    pub fn shadow(mut self, s: ShadowProps) -> Self { self.opts.shadow = Some(s); self }
    /// Set the corner radius in inches for rounded rectangles.
    pub fn rect_radius(mut self, r: f64) -> Self { self.opts.rect_radius = Some(r); self }
    /// Attach a hyperlink activated on click.
    pub fn hyperlink(mut self, h: HyperlinkProps) -> Self { self.opts.hyperlink = Some(h); self }
    /// Attach a hyperlink activated on hover.
    pub fn hover(mut self, h: HyperlinkProps) -> Self { self.opts.hover = Some(h); self }
    /// Set the accessibility alt text for the shape.
    pub fn alt_text(mut self, t: impl Into<String>) -> Self { self.opts.alt_text = Some(t.into()); self }
    /// Set angle range for PIE / ARC / BLOCK_ARC shapes.
    /// `start` and `swing` are in degrees (clockwise from east/3 o'clock).
    pub fn angle_range(mut self, start: f64, swing: f64) -> Self {
        self.opts.angle_range = Some([start, swing]);
        self
    }
    /// Set inner-radius ratio 0.0–1.0 for BLOCK_ARC (default ~0.5 if omitted).
    pub fn arc_thickness(mut self, ratio: f64) -> Self { self.opts.arc_thickness = Some(ratio); self }
    /// Set custom freeform geometry. The shape_type is ignored when this is set.
    pub fn custom_geometry(mut self, pts: Vec<CustomGeomPoint>) -> Self {
        self.opts.custom_geometry = Some(pts);
        self
    }
    /// Apply a pattern fill to the shape.
    pub fn pattern_fill(mut self, p: PatternFill) -> Self {
        self.opts.fill = Some(ShapeFillProps {
            fill_type: crate::types::FillType::Pattern,
            color: None,
            transparency: None,
            gradient: None,
            pattern: Some(p),
        });
        self
    }
    /// Add a click-triggered animation effect.
    pub fn animation(mut self, anim: AnimationEffect) -> Self { self.opts.animations.push(anim); self }

    /// Apply 3D shape effects (bevel, extrusion, material).
    pub fn shape_3d(mut self, props: Shape3DProps) -> Self { self.opts.shape_3d = Some(props); self }
    /// Apply a 3D scene (camera and light rig).
    pub fn scene_3d(mut self, props: Scene3DProps) -> Self { self.opts.scene_3d = Some(props); self }

    /// Consume the builder and return the configured shape options.
    pub fn build(self) -> ShapeOptions {
        self.opts
    }
}

impl Default for ShapeOptionsBuilder {
    fn default() -> Self {
        Self::new()
    }
}
