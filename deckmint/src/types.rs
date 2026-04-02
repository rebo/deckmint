use crate::enums::SchemeColor;

/// A coordinate value — either inches, EMU (>= 100), or a percentage ("75%")
#[derive(Debug, Clone)]
pub enum Coord {
    /// Inches (< 100)
    Inches(f64),
    /// Already in EMU (>= 100)
    Emu(i64),
    /// Percentage of slide dimension, 0.0–100.0
    Percent(f64),
}

impl Coord {
    /// Convert to EMU given the layout dimension in EMU for percentage resolution.
    pub fn to_emu(&self, layout_dim: i64) -> i64 {
        use crate::enums::EMU;
        match self {
            Coord::Inches(in_val) => {
                if *in_val > 100.0 {
                    *in_val as i64 // already EMU
                } else {
                    (EMU as f64 * in_val).round() as i64
                }
            }
            Coord::Emu(v) => *v,
            Coord::Percent(pct) => ((pct / 100.0) * layout_dim as f64).round() as i64,
        }
    }
}

/// Luminance/tint/shade modifiers for a theme color.
/// Values are in OOXML thousandths-of-a-percent (e.g. 75000 = 75%).
#[derive(Debug, Clone, Default)]
pub struct ThemeColorMod {
    /// Luminance multiply, e.g. 75000 = 75%
    pub lum_mod: Option<i32>,
    /// Luminance offset, e.g. 15000 = 15%
    pub lum_off: Option<i32>,
    /// Tint (lighten toward white), e.g. 50000 = 50%
    pub tint: Option<i32>,
    /// Shade (darken toward black), e.g. 50000 = 50%
    pub shade: Option<i32>,
    /// Saturation multiply
    pub sat_mod: Option<i32>,
}

impl ThemeColorMod {
    /// Create a modifier with luminance multiply and offset values.
    pub fn lum(lum_mod: i32, lum_off: i32) -> Self {
        ThemeColorMod { lum_mod: Some(lum_mod), lum_off: Some(lum_off), ..Default::default() }
    }
    /// Create a modifier that tints toward white by the given amount.
    pub fn tint(t: i32) -> Self {
        ThemeColorMod { tint: Some(t), ..Default::default() }
    }
    /// Create a modifier that shades toward black by the given amount.
    pub fn shade(s: i32) -> Self {
        ThemeColorMod { shade: Some(s), ..Default::default() }
    }
}

/// A color value — either a 6-digit hex string or a scheme color reference
#[derive(Debug, Clone)]
pub enum Color {
    /// Literal 6-digit hex colour string (no `#` prefix).
    Hex(String),
    /// A reference to a theme/scheme colour.
    Theme(SchemeColor),
    /// Theme color with luminance/tint/shade modifiers.
    ThemedWith(SchemeColor, ThemeColorMod),
}

impl Color {
    /// Parse from a string: tries hex first, then scheme color names
    pub fn from_str(s: &str) -> Self {
        let clean = s.trim_start_matches('#');
        if crate::enums::is_hex_color(clean) {
            Color::Hex(clean.to_uppercase())
        } else if let Some(sc) = SchemeColor::from_str(clean) {
            Color::Theme(sc)
        } else {
            Color::Hex(crate::enums::DEF_FONT_COLOR.to_string())
        }
    }

    /// Returns the OOXML color element inner value string
    pub fn as_ooxml_val(&self) -> String {
        match self {
            Color::Hex(h) => h.to_uppercase(),
            Color::Theme(sc) => sc.as_str().to_string(),
            Color::ThemedWith(sc, _) => sc.as_str().to_string(),
        }
    }

    /// Return `true` if this is a literal hex colour (not a theme reference).
    pub fn is_hex(&self) -> bool {
        matches!(self, Color::Hex(_))
    }
}

/// A single drawing operation in a custom freeform shape path.
///
/// Coordinates are fractions of the shape's bounding box: `(0.0, 0.0)` = top-left,
/// `(1.0, 1.0)` = bottom-right.
///
/// # Example
/// ```rust,no_run
/// use deckmint::types::CustomGeomPoint as CGP;
/// // Triangle
/// let pts = vec![
///     CGP::MoveTo(0.5, 0.0),   // top-center
///     CGP::LineTo(1.0, 1.0),   // bottom-right
///     CGP::LineTo(0.0, 1.0),   // bottom-left
///     CGP::Close,
/// ];
/// ```
#[derive(Debug, Clone)]
pub enum CustomGeomPoint {
    /// Start of a sub-path at `(x, y)`.
    MoveTo(f64, f64),
    /// Straight line to `(x, y)`.
    LineTo(f64, f64),
    /// Arc whose ellipse has radii `(w_r, h_r)` (fractions of shape size),
    /// starting at `start_angle` degrees and sweeping `swing_angle` degrees clockwise.
    ArcTo { w_r: f64, h_r: f64, start_angle: f64, swing_angle: f64 },
    /// Cubic Bézier to `end` via control points `cp1` and `cp2`.
    /// Fields: (cp1_x, cp1_y, cp2_x, cp2_y, end_x, end_y)
    CubicBezTo(f64, f64, f64, f64, f64, f64),
    /// Quadratic Bézier to `end` via control point `cp`.
    /// Fields: (cp_x, cp_y, end_x, end_y)
    QuadBezTo(f64, f64, f64, f64),
    /// Close the current sub-path (draws a straight line back to the last MoveTo).
    Close,
}

/// Margin — either uniform or TRBL (top/right/bottom/left)
#[derive(Debug, Clone)]
pub enum Margin {
    /// Same margin on all four sides, in inches.
    Uniform(f64),
    /// Per-side margins as `[top, right, bottom, left]` in inches.
    Trbl([f64; 4]),
}

impl Default for Margin {
    fn default() -> Self {
        Margin::Trbl(crate::enums::DEF_CELL_MARGIN_IN)
    }
}

impl Margin {
    /// Top margin in inches.
    pub fn top(&self) -> f64 {
        match self {
            Margin::Uniform(v) => *v,
            Margin::Trbl(a) => a[0],
        }
    }
    /// Right margin in inches.
    pub fn right(&self) -> f64 {
        match self {
            Margin::Uniform(v) => *v,
            Margin::Trbl(a) => a[1],
        }
    }
    /// Bottom margin in inches.
    pub fn bottom(&self) -> f64 {
        match self {
            Margin::Uniform(v) => *v,
            Margin::Trbl(a) => a[2],
        }
    }
    /// Left margin in inches.
    pub fn left(&self) -> f64 {
        match self {
            Margin::Uniform(v) => *v,
            Margin::Trbl(a) => a[3],
        }
    }
}

impl From<f64> for Margin {
    /// Create a uniform margin from a single value in inches.
    fn from(v: f64) -> Self {
        Margin::Uniform(v)
    }
}

impl From<[f64; 4]> for Margin {
    /// Create a TRBL margin from `[top, right, bottom, left]` in inches.
    fn from(arr: [f64; 4]) -> Self {
        Margin::Trbl(arr)
    }
}

/// Position and size properties shared by all slide objects
#[derive(Debug, Clone, Default)]
pub struct PositionProps {
    /// Horizontal offset from the left edge.
    pub x: Option<Coord>,
    /// Vertical offset from the top edge.
    pub y: Option<Coord>,
    /// Width.
    pub w: Option<Coord>,
    /// Height.
    pub h: Option<Coord>,
}

/// Presentation layout (slide dimensions)
#[derive(Debug, Clone)]
pub struct PresLayout {
    pub name: String,
    /// Width in EMU
    pub width: i64,
    /// Height in EMU
    pub height: i64,
}

impl PresLayout {
    /// Standard 16:9 widescreen layout (10" × 5.63").
    pub fn layout_16x9() -> Self {
        PresLayout { name: "screen16x9".to_string(), width: 9_144_000, height: 5_143_500 }
    }
    /// Standard 4:3 layout (10" × 7.5").
    pub fn layout_4x3() -> Self {
        PresLayout { name: "screen4x3".to_string(), width: 9_144_000, height: 6_858_000 }
    }
    /// 16:10 widescreen layout (10" × 6.25").
    pub fn layout_16x10() -> Self {
        PresLayout { name: "screen16x10".to_string(), width: 9_144_000, height: 5_715_000 }
    }
    /// Extra-wide custom layout (13.33" × 7.5").
    pub fn layout_wide() -> Self {
        PresLayout { name: "custom".to_string(), width: 12_192_000, height: 6_858_000 }
    }
}

impl Default for PresLayout {
    fn default() -> Self {
        PresLayout::layout_16x9()
    }
}

/// Border properties
#[derive(Debug, Clone)]
pub struct BorderProps {
    /// Line style.
    pub border_type: BorderType,
    /// Hex color, no `#` prefix (e.g. `"666666"`).
    pub color: Option<String>,
    /// Line thickness in points.
    pub pt: f64,
}

/// Border line style.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum BorderType {
    /// Solid line.
    #[default]
    Solid,
    /// Dashed line.
    Dash,
    /// No border.
    None,
}

impl Default for BorderProps {
    fn default() -> Self {
        BorderProps { border_type: BorderType::Solid, color: Some("666666".to_string()), pt: 1.0 }
    }
}

/// Shadow properties
#[derive(Debug, Clone)]
pub struct ShadowProps {
    /// Outer or inner shadow.
    pub shadow_type: ShadowType,
    /// Opacity, 0.0 (transparent) – 1.0 (opaque).
    pub opacity: Option<f64>,
    /// Blur radius in points.
    pub blur: Option<f64>,
    /// Distance from the shape in points.
    pub offset: Option<f64>,
    /// Light-source angle in degrees (0 = right, 90 = bottom, etc.).
    pub angle: Option<f64>,
    /// Shadow hex color, no `#` prefix.
    pub color: Option<String>,
    /// Whether the shadow rotates with the shape.
    pub rotate_with_shape: bool,
}

/// Shadow direction.
#[derive(Debug, Clone, PartialEq)]
pub enum ShadowType {
    /// Drop shadow outside the shape.
    Outer,
    /// Inset shadow inside the shape.
    Inner,
    /// No shadow.
    None,
}

impl Default for ShadowProps {
    fn default() -> Self {
        ShadowProps {
            shadow_type: ShadowType::Outer,
            blur: Some(3.0),
            offset: Some(23000.0 / 12700.0),
            angle: Some(90.0),
            color: Some("000000".to_string()),
            opacity: Some(0.35),
            rotate_with_shape: true,
        }
    }
}

impl ShadowProps {
    /// Create a standard outer drop shadow with sensible defaults.
    pub fn outer() -> Self {
        ShadowProps::default()
    }

    /// Create an inner (inset) shadow with sensible defaults.
    pub fn inner() -> Self {
        ShadowProps { shadow_type: ShadowType::Inner, ..ShadowProps::default() }
    }

    /// Set the shadow hex color, no `#` prefix.
    pub fn with_color(mut self, c: impl Into<String>) -> Self {
        self.color = Some(c.into().trim_start_matches('#').to_uppercase());
        self
    }

    /// Set the blur radius in points.
    pub fn with_blur(mut self, pt: f64) -> Self {
        self.blur = Some(pt);
        self
    }

    /// Set the shadow offset distance in points.
    pub fn with_offset(mut self, pt: f64) -> Self {
        self.offset = Some(pt);
        self
    }

    /// Set the light-source angle in degrees.
    pub fn with_angle(mut self, deg: f64) -> Self {
        self.angle = Some(deg);
        self
    }

    /// Set the shadow opacity, 0.0 (transparent) – 1.0 (opaque).
    pub fn with_opacity(mut self, v: f64) -> Self {
        self.opacity = Some(v);
        self
    }
}

// ─────────────────────────────────────────────────────────────
// 3D effects
// ─────────────────────────────────────────────────────────────

/// Bevel preset type for 3D shape effects.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BevelPreset {
    Circle,
    RelaxedInset,
    Angle,
    SoftRound,
    Convex,
    Slope,
    Divot,
    Riblet,
    HardEdge,
    ArtDeco,
    Cross,
    CoolSlant,
}

impl BevelPreset {
    /// OOXML `prst` attribute value.
    pub fn as_str(&self) -> &'static str {
        match self {
            BevelPreset::Circle => "circle",
            BevelPreset::RelaxedInset => "relaxedInset",
            BevelPreset::Angle => "angle",
            BevelPreset::SoftRound => "softRound",
            BevelPreset::Convex => "convex",
            BevelPreset::Slope => "slope",
            BevelPreset::Divot => "divot",
            BevelPreset::Riblet => "riblet",
            BevelPreset::HardEdge => "hardEdge",
            BevelPreset::ArtDeco => "artDeco",
            BevelPreset::Cross => "cross",
            BevelPreset::CoolSlant => "coolSlant",
        }
    }
}

/// Bevel properties for top or bottom bevel.
#[derive(Debug, Clone)]
pub struct BevelProps {
    /// Bevel preset shape.
    pub preset: BevelPreset,
    /// Width in EMU (default 76200 = 6pt).
    pub width: Option<i64>,
    /// Height in EMU (default 76200 = 6pt).
    pub height: Option<i64>,
}

impl BevelProps {
    /// Create a bevel with default width/height.
    pub fn new(preset: BevelPreset) -> Self {
        BevelProps { preset, width: None, height: None }
    }
    /// Set width and height in EMU.
    pub fn with_size(mut self, w: i64, h: i64) -> Self {
        self.width = Some(w);
        self.height = Some(h);
        self
    }
}

/// Material preset for 3D surfaces.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum MaterialPreset {
    Matte,
    WarmMatte,
    Plastic,
    Metal,
    DarkEdge,
    SoftEdge,
    Flat,
    SoftMetal,
    Powder,
    TranslucentPowder,
    Clear,
}

impl MaterialPreset {
    pub fn as_str(&self) -> &'static str {
        match self {
            MaterialPreset::Matte => "matte",
            MaterialPreset::WarmMatte => "warmMatte",
            MaterialPreset::Plastic => "plastic",
            MaterialPreset::Metal => "metal",
            MaterialPreset::DarkEdge => "dkEdge",
            MaterialPreset::SoftEdge => "softEdge",
            MaterialPreset::Flat => "flat",
            MaterialPreset::SoftMetal => "softmetal",
            MaterialPreset::Powder => "powder",
            MaterialPreset::TranslucentPowder => "translucentPowder",
            MaterialPreset::Clear => "clear",
        }
    }
}

/// Camera preset for 3D scene.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CameraPreset {
    OrthographicFront,
    PerspectiveFront,
    IsometricTopUp,
    IsometricTopDown,
    IsometricLeftDown,
    IsometricRightUp,
    ObliqueTopLeft,
    ObliqueTop,
    ObliqueTopRight,
}

impl CameraPreset {
    pub fn as_str(&self) -> &'static str {
        match self {
            CameraPreset::OrthographicFront => "orthographicFront",
            CameraPreset::PerspectiveFront => "perspectiveFront",
            CameraPreset::IsometricTopUp => "isometricTopUp",
            CameraPreset::IsometricTopDown => "isometricTopDown",
            CameraPreset::IsometricLeftDown => "isometricLeftDown",
            CameraPreset::IsometricRightUp => "isometricRightUp",
            CameraPreset::ObliqueTopLeft => "obliqueTopLeft",
            CameraPreset::ObliqueTop => "obliqueTop",
            CameraPreset::ObliqueTopRight => "obliqueTopRight",
        }
    }
}

/// Light rig type for 3D scene.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LightRigType {
    ThreePt,
    Balanced,
    Harsh,
    Flood,
    Flat,
    Soft,
    Morning,
    Sunrise,
    Sunset,
    Chilly,
    Freezing,
    Glow,
    BrightRoom,
    TwoPt,
    Contrasting,
}

impl LightRigType {
    pub fn as_str(&self) -> &'static str {
        match self {
            LightRigType::ThreePt => "threePt",
            LightRigType::Balanced => "balanced",
            LightRigType::Harsh => "harsh",
            LightRigType::Flood => "flood",
            LightRigType::Flat => "flat",
            LightRigType::Soft => "soft",
            LightRigType::Morning => "morning",
            LightRigType::Sunrise => "sunrise",
            LightRigType::Sunset => "sunset",
            LightRigType::Chilly => "chilly",
            LightRigType::Freezing => "freezing",
            LightRigType::Glow => "glow",
            LightRigType::BrightRoom => "brightRoom",
            LightRigType::TwoPt => "twoPt",
            LightRigType::Contrasting => "contrasting",
        }
    }
}

/// Light direction in a 3D scene.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LightDirection {
    Top,
    TopLeft,
    TopRight,
    Left,
    Right,
    Bottom,
    BottomLeft,
    BottomRight,
}

impl LightDirection {
    pub fn as_str(&self) -> &'static str {
        match self {
            LightDirection::Top => "t",
            LightDirection::TopLeft => "tl",
            LightDirection::TopRight => "tr",
            LightDirection::Left => "l",
            LightDirection::Right => "r",
            LightDirection::Bottom => "b",
            LightDirection::BottomLeft => "bl",
            LightDirection::BottomRight => "br",
        }
    }
}

/// 3D rotation angles, in 60,000ths of a degree.
#[derive(Debug, Clone, Copy)]
pub struct Rotation3D {
    /// Latitude (x-axis rotation).
    pub lat: i64,
    /// Longitude (y-axis rotation).
    pub lon: i64,
    /// Revolution (z-axis rotation).
    pub rev: i64,
}

impl Rotation3D {
    /// Create from degree values (automatically converts to 60,000ths).
    pub fn from_degrees(lat: f64, lon: f64, rev: f64) -> Self {
        Rotation3D {
            lat: (lat * 60_000.0).round() as i64,
            lon: (lon * 60_000.0).round() as i64,
            rev: (rev * 60_000.0).round() as i64,
        }
    }
}

/// 3D shape properties (bevel, extrusion, material).
#[derive(Debug, Clone)]
pub struct Shape3DProps {
    /// Top bevel effect.
    pub bevel_top: Option<BevelProps>,
    /// Bottom bevel effect.
    pub bevel_bottom: Option<BevelProps>,
    /// Extrusion depth in EMU.
    pub extrusion_height: Option<i64>,
    /// Contour width in EMU.
    pub contour_width: Option<i64>,
    /// Contour colour (hex, no `#` prefix).
    pub contour_color: Option<String>,
    /// Surface material preset.
    pub material: Option<MaterialPreset>,
}

impl Default for Shape3DProps {
    fn default() -> Self {
        Shape3DProps {
            bevel_top: None, bevel_bottom: None,
            extrusion_height: None, contour_width: None,
            contour_color: None, material: None,
        }
    }
}

/// Camera properties for a 3D scene.
#[derive(Debug, Clone)]
pub struct Camera3D {
    /// Camera preset position.
    pub preset: CameraPreset,
    /// Field of view in 60,000ths of a degree (optional).
    pub fov: Option<i64>,
    /// Camera rotation (optional).
    pub rotation: Option<Rotation3D>,
}

/// Light rig properties for a 3D scene.
#[derive(Debug, Clone)]
pub struct LightRig3D {
    /// Light rig type.
    pub rig_type: LightRigType,
    /// Light direction.
    pub direction: LightDirection,
    /// Light rotation (optional).
    pub rotation: Option<Rotation3D>,
}

/// 3D scene properties (camera + light rig).
#[derive(Debug, Clone)]
pub struct Scene3DProps {
    /// Camera settings.
    pub camera: Camera3D,
    /// Light rig settings.
    pub light_rig: LightRig3D,
}

/// A single stop in a gradient fill.
#[derive(Debug, Clone)]
pub struct GradientStop {
    /// Color for this stop — supports hex, theme, and themed-with-modifiers.
    pub color: Color,
    /// Position along the gradient, 0.0 (start) – 100.0 (end).
    pub position: f64,
    /// Optional transparency, 0.0 (opaque) – 100.0 (fully transparent).
    pub transparency: Option<f64>,
}

impl GradientStop {
    /// Create a gradient stop from a hex colour string and position (0.0–100.0).
    pub fn new(color: impl Into<String>, position: f64) -> Self {
        GradientStop { color: Color::Hex(color.into().trim_start_matches('#').to_uppercase()), position, transparency: None }
    }
    /// Create a stop from a `Color` value (supports theme colors).
    pub fn from_color(color: Color, position: f64) -> Self {
        GradientStop { color, position, transparency: None }
    }
    /// Set the transparency for this stop, 0.0 (opaque) – 100.0 (fully transparent).
    pub fn with_transparency(mut self, t: f64) -> Self { self.transparency = Some(t); self }
}

/// Whether the gradient radiates linearly along an angle or circularly from the centre.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum GradientType {
    /// Linear gradient along an angle.
    #[default]
    Linear,
    /// Radial (circular) gradient from centre outward.
    Radial,
}

/// A gradient fill definition.
#[derive(Debug, Clone)]
pub struct GradientFill {
    /// Colour stops — supply at least two.
    pub stops: Vec<GradientStop>,
    /// Gradient angle in degrees (linear only).
    /// `0` = left→right, `90` = top→bottom, `180` = right→left, `270` = bottom→top.
    pub angle: f64,
    /// Whether the gradient is linear or radial.
    pub gradient_type: GradientType,
}

impl GradientFill {
    /// Two-stop linear gradient between `from` and `to` at the given angle.
    pub fn two_color(angle: f64, from: impl Into<String>, to: impl Into<String>) -> Self {
        GradientFill {
            stops: vec![GradientStop::new(from, 0.0), GradientStop::new(to, 100.0)],
            angle,
            gradient_type: GradientType::Linear,
        }
    }

    /// Linear gradient with full stop control.
    pub fn linear(angle: f64, stops: Vec<GradientStop>) -> Self {
        GradientFill { stops, angle, gradient_type: GradientType::Linear }
    }

    /// Radial (circular) gradient — angle is ignored.
    pub fn radial(stops: Vec<GradientStop>) -> Self {
        GradientFill { stops, angle: 0.0, gradient_type: GradientType::Radial }
    }

    /// Two-stop radial gradient (inner colour first, outer colour last).
    pub fn radial_two_color(inner: impl Into<String>, outer: impl Into<String>) -> Self {
        GradientFill::radial(vec![
            GradientStop::new(inner, 0.0),
            GradientStop::new(outer, 100.0),
        ])
    }
}

/// OOXML preset pattern types for pattern fills
#[derive(Debug, Clone, PartialEq)]
pub enum PatternType {
    /// Cross pattern.
    Cross,
    /// Dark downward diagonal pattern.
    DarkDnDiag,
    /// Dark horizontal pattern.
    DarkHorz,
    /// Dark upward diagonal pattern.
    DarkUpDiag,
    /// Dark vertical pattern.
    DarkVert,
    /// Downward diagonal pattern.
    DnDiag,
    /// Dotted diamond pattern.
    DotDmnd,
    /// Dotted grid pattern.
    DotGrid,
    /// Horizontal pattern.
    Horz,
    /// Horizontal brick pattern.
    HorzBrick,
    /// Large checkerboard pattern.
    LgCheck,
    /// Large confetti pattern.
    LgConfetti,
    /// Large grid pattern.
    LgGrid,
    /// Light downward diagonal pattern.
    LtDnDiag,
    /// Light horizontal pattern.
    LtHorz,
    /// Light upward diagonal pattern.
    LtUpDiag,
    /// Light vertical pattern.
    LtVert,
    /// Narrow horizontal pattern.
    NarHorz,
    /// Narrow vertical pattern.
    NarVert,
    /// Open diamond pattern.
    OpenDmnd,
    /// 5% fill pattern.
    Pct5,
    /// 10% fill pattern.
    Pct10,
    /// 20% fill pattern.
    Pct20,
    /// 25% fill pattern.
    Pct25,
    /// 30% fill pattern.
    Pct30,
    /// 40% fill pattern.
    Pct40,
    /// 50% fill pattern.
    Pct50,
    /// 60% fill pattern.
    Pct60,
    /// 70% fill pattern.
    Pct70,
    /// 75% fill pattern.
    Pct75,
    /// 80% fill pattern.
    Pct80,
    /// 90% fill pattern.
    Pct90,
    /// Shingle pattern.
    Shingle,
    /// Small checkerboard pattern.
    SmCheck,
    /// Small confetti pattern.
    SmConfetti,
    /// Small grid pattern.
    SmGrid,
    /// Solid diamond pattern.
    SolidDmnd,
    /// Sphere pattern.
    Sphere,
    /// Trellis pattern.
    Trellis,
    /// Upward diagonal pattern.
    UpDiag,
    /// Vertical pattern.
    Vert,
    /// Wave pattern.
    Wave,
    /// Wide downward diagonal pattern.
    WdDnDiag,
    /// Wide upward diagonal pattern.
    WdUpDiag,
    /// Zig-zag pattern.
    ZigZag,
}

impl PatternType {
    /// Return the OOXML preset pattern name string.
    pub fn as_str(&self) -> &'static str {
        match self {
            PatternType::Cross => "cross",
            PatternType::DarkDnDiag => "darkDnDiag",
            PatternType::DarkHorz => "darkHorz",
            PatternType::DarkUpDiag => "darkUpDiag",
            PatternType::DarkVert => "darkVert",
            PatternType::DnDiag => "dnDiag",
            PatternType::DotDmnd => "dotDmnd",
            PatternType::DotGrid => "dotGrid",
            PatternType::Horz => "horz",
            PatternType::HorzBrick => "horzBrick",
            PatternType::LgCheck => "lgCheck",
            PatternType::LgConfetti => "lgConfetti",
            PatternType::LgGrid => "lgGrid",
            PatternType::LtDnDiag => "ltDnDiag",
            PatternType::LtHorz => "ltHorz",
            PatternType::LtUpDiag => "ltUpDiag",
            PatternType::LtVert => "ltVert",
            PatternType::NarHorz => "narHorz",
            PatternType::NarVert => "narVert",
            PatternType::OpenDmnd => "openDmnd",
            PatternType::Pct5 => "pct5",
            PatternType::Pct10 => "pct10",
            PatternType::Pct20 => "pct20",
            PatternType::Pct25 => "pct25",
            PatternType::Pct30 => "pct30",
            PatternType::Pct40 => "pct40",
            PatternType::Pct50 => "pct50",
            PatternType::Pct60 => "pct60",
            PatternType::Pct70 => "pct70",
            PatternType::Pct75 => "pct75",
            PatternType::Pct80 => "pct80",
            PatternType::Pct90 => "pct90",
            PatternType::Shingle => "shingle",
            PatternType::SmCheck => "smCheck",
            PatternType::SmConfetti => "smConfetti",
            PatternType::SmGrid => "smGrid",
            PatternType::SolidDmnd => "solidDmnd",
            PatternType::Sphere => "sphere",
            PatternType::Trellis => "trellis",
            PatternType::UpDiag => "upDiag",
            PatternType::Vert => "vert",
            PatternType::Wave => "wave",
            PatternType::WdDnDiag => "wdDnDiag",
            PatternType::WdUpDiag => "wdUpDiag",
            PatternType::ZigZag => "zigZag",
        }
    }
}

/// Pattern fill — foreground + background color with an OOXML preset pattern
#[derive(Debug, Clone)]
pub struct PatternFill {
    /// The preset pattern type.
    pub pattern: PatternType,
    /// Foreground hex color (no `#` prefix).
    pub fg_color: String,
    /// Background hex color (no `#` prefix).
    pub bg_color: String,
}

/// Fill properties for shapes
#[derive(Debug, Clone)]
pub struct ShapeFillProps {
    /// Solid, gradient, or no fill.
    pub fill_type: FillType,
    /// Color for solid fills.
    pub color: Option<Color>,
    /// Transparency, 0.0 (opaque) – 100.0 (fully transparent).
    pub transparency: Option<f64>,
    /// Gradient definition (used when `fill_type` is [`FillType::Gradient`]).
    pub gradient: Option<GradientFill>,
    /// Pattern fill definition (used when `fill_type` is [`FillType::Pattern`]).
    pub pattern: Option<PatternFill>,
}

/// Fill style for shapes and table cells.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum FillType {
    /// Solid single-colour fill.
    #[default]
    Solid,
    /// Multi-stop gradient fill.
    Gradient,
    /// Pattern fill.
    Pattern,
    /// Transparent / no fill.
    None,
}

impl Default for ShapeFillProps {
    fn default() -> Self {
        ShapeFillProps { fill_type: FillType::Solid, color: None, transparency: None, gradient: None, pattern: None }
    }
}

/// Line cap style for shape outlines.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LineCap {
    /// Flat cap (no extension beyond the endpoint).
    Flat,
    /// Round cap (semicircle at endpoint).
    Round,
    /// Square cap (half-width extension beyond endpoint).
    Square,
}

impl LineCap {
    /// OOXML attribute value for the `cap` attribute on `<a:ln>`.
    pub fn as_str(&self) -> &'static str {
        match self {
            LineCap::Flat => "flat",
            LineCap::Round => "rnd",
            LineCap::Square => "sq",
        }
    }
}

/// Line join style for shape outlines.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum LineJoin {
    /// Round join (arc at corner).
    Round,
    /// Bevel join (straight cut at corner).
    Bevel,
    /// Miter join (sharp point at corner).
    Miter,
}

/// Line properties for shapes
#[derive(Debug, Clone, Default)]
pub struct ShapeLineProps {
    /// Line hex color, no `#` prefix.
    pub color: Option<String>,
    /// Transparency, 0.0 (opaque) – 100.0 (fully transparent).
    pub transparency: Option<f64>,
    /// Line width in points.
    pub width: Option<f64>,
    /// OOXML dash style (e.g. `"dash"`, `"lgDash"`, `"sysDot"`).
    pub dash_type: Option<String>,
    /// Start arrow type (e.g. `"triangle"`, `"arrow"`).
    pub begin_arrow_type: Option<String>,
    /// End arrow type (e.g. `"triangle"`, `"arrow"`).
    pub end_arrow_type: Option<String>,
    /// Line cap style (flat, round, or square).
    pub cap: Option<LineCap>,
    /// Line join style (round, bevel, or miter).
    pub join: Option<LineJoin>,
}

/// Slide number properties
#[derive(Debug, Clone, Default)]
pub struct SlideNumberProps {
    /// Horizontal offset from the left edge.
    pub x: Option<Coord>,
    /// Vertical offset from the top edge.
    pub y: Option<Coord>,
    /// Width of the slide-number placeholder.
    pub w: Option<Coord>,
    /// Height of the slide-number placeholder.
    pub h: Option<Coord>,
    /// Font family name (e.g. `"Arial"`).
    pub font_face: Option<String>,
    /// Font size in points.
    pub font_size: Option<f64>,
    /// Hex color, no `#` prefix.
    pub color: Option<String>,
    /// Whether the number is bold.
    pub bold: bool,
    /// Horizontal alignment (e.g. `"center"`, `"right"`).
    pub align: Option<String>,
    /// Vertical alignment (e.g. `"middle"`, `"bottom"`).
    pub valign: Option<String>,
    /// Inner margin of the placeholder.
    pub margin: Option<Margin>,
}

/// Navigation action types for hyperlinks — these require no relationship ID.
#[derive(Debug, Clone, PartialEq)]
pub enum HyperlinkAction {
    /// Navigate to the next slide.
    NextSlide,
    /// Navigate to the previous slide.
    PrevSlide,
    /// Navigate to the first slide.
    FirstSlide,
    /// Navigate to the last slide.
    LastSlide,
    /// End the slide show.
    EndShow,
}

impl HyperlinkAction {
    /// Return the OOXML `ppaction://` URI for this navigation action.
    pub fn as_ppaction(&self) -> &'static str {
        match self {
            HyperlinkAction::NextSlide  => "ppaction://hlinkshowjump?jump=nextslide",
            HyperlinkAction::PrevSlide  => "ppaction://hlinkshowjump?jump=previousslide",
            HyperlinkAction::FirstSlide => "ppaction://hlinkshowjump?jump=firstslide",
            HyperlinkAction::LastSlide  => "ppaction://hlinkshowjump?jump=lastslide",
            HyperlinkAction::EndShow    => "ppaction://hlinkshowjump?jump=endshow",
        }
    }
}

/// Hyperlink properties
#[derive(Debug, Clone, Default)]
pub struct HyperlinkProps {
    /// Relationship ID (assigned internally when adding objects).
    pub r_id: u32,
    /// Target slide number for an internal slide-jump link.
    pub slide: Option<u32>,
    /// External URL (e.g. `"https://example.com"`).
    pub url: Option<String>,
    /// Hover tooltip text.
    pub tooltip: Option<String>,
    /// Navigation action (NextSlide, PrevSlide, etc.) — takes priority over url/slide.
    /// No relationship is registered when this is set.
    pub action: Option<HyperlinkAction>,
}

impl HyperlinkProps {
    /// Create a hyperlink to an external URL.
    pub fn url(url: impl Into<String>) -> Self {
        HyperlinkProps { r_id: 0, slide: None, url: Some(url.into()), tooltip: None, action: None }
    }

    /// Create a hyperlink that jumps to a slide by number (1-based).
    pub fn slide(slide_num: u32) -> Self {
        HyperlinkProps { r_id: 0, slide: Some(slide_num), url: None, tooltip: None, action: None }
    }

    /// Set the hover tooltip text.
    pub fn with_tooltip(mut self, tip: impl Into<String>) -> Self {
        self.tooltip = Some(tip.into());
        self
    }
}

/// Glow effect for text runs
#[derive(Debug, Clone)]
pub struct GlowProps {
    /// Glow radius in points
    pub size: f64,
    /// Hex color string (no #)
    pub color: String,
    /// Opacity 0.0–1.0
    pub opacity: f64,
}

impl GlowProps {
    /// Create a glow effect with the given radius (pt), hex color, and opacity (0.0–1.0).
    pub fn new(size: f64, color: impl Into<String>, opacity: f64) -> Self {
        GlowProps { size, color: color.into().trim_start_matches('#').to_uppercase(), opacity }
    }
}

/// Text outline (stroke) for text runs
#[derive(Debug, Clone)]
pub struct TextOutlineProps {
    /// Hex color string (no #)
    pub color: String,
    /// Outline width in points
    pub size: f64,
}

// ─── Animation ───────────────────────────────────────────────

/// Direction used by directional animation effects (FlyIn, WipeIn, etc.).
#[derive(Debug, Clone, PartialEq)]
pub enum Direction {
    /// From / toward the left edge.
    Left,
    /// From / toward the right edge.
    Right,
    /// From / toward the top edge.
    Up,
    /// From / toward the bottom edge.
    Down,
}

/// Orientation for Split (barn-door) animations.
#[derive(Debug, Clone, PartialEq)]
pub enum SplitOrientation {
    /// Horizontal split (top and bottom halves).
    Horizontal,
    /// Vertical split (left and right halves).
    Vertical,
}

/// Direction pattern for Checkerboard animations.
#[derive(Debug, Clone, PartialEq)]
pub enum CheckerboardDir {
    /// Checkerboard fills across (left→right).
    Across,
    /// Checkerboard fills downward (top→bottom).
    Down,
}

/// Direction for diagonal Strips animations.
#[derive(Debug, Clone, PartialEq)]
pub enum StripDir {
    /// Diagonal strips from top-right to bottom-left.
    LeftDown,
    /// Diagonal strips from bottom-right to top-left.
    LeftUp,
    /// Diagonal strips from top-left to bottom-right.
    RightDown,
    /// Diagonal strips from bottom-left to top-right.
    RightUp,
}

/// Shape variant for Shape (Box/Circle/Diamond/Plus) entrance animations.
#[derive(Debug, Clone, PartialEq)]
pub enum ShapeVariant {
    /// Rectangular shape.
    Box,
    /// Circular shape.
    Circle,
    /// Diamond (rotated square) shape.
    Diamond,
    /// Plus / cross shape.
    Plus,
}

/// Colour adjustments applied to an image at render time.
#[derive(Debug, Clone, Default)]
pub struct ImageColorAdjust {
    /// Brightness adjustment, −100.0 (fully dark) to +100.0 (fully bright). 0.0 = no change.
    pub brightness: Option<f64>,
    /// Contrast adjustment, −100.0 to +100.0. 0.0 = no change.
    pub contrast: Option<f64>,
    /// Convert image to greyscale.
    pub grayscale: bool,
}

/// An auto-updating field that can be placed in a text run.
#[derive(Debug, Clone)]
pub enum FieldType {
    /// Shows the current slide number.
    SlideNumber,
    /// Auto-updating date/time (format determined by PowerPoint locale).
    DateTime,
    /// Raw OOXML field type string (e.g. `"datetime1"`, `"slidenum"`).
    Custom(String),
}

impl FieldType {
    /// Return the OOXML field type string for XML output.
    pub fn as_ooxml(&self) -> &str {
        match self {
            FieldType::SlideNumber => "slidenum",
            FieldType::DateTime => "datetime1",
            FieldType::Custom(s) => s.as_str(),
        }
    }
}

/// Targets a sub-range of text within a shape for an animation.
/// Indices are zero-based and inclusive on both ends.
#[derive(Debug, Clone, PartialEq)]
pub enum TextTarget {
    /// Animate only the characters from `st_idx` to `end_idx` (inclusive).
    CharRange { st_idx: u32, end_idx: u32 },
    /// Animate only the paragraphs from `st_idx` to `end_idx` (inclusive).
    ParaRange { st_idx: u32, end_idx: u32 },
}

/// When an animation fires relative to the previous animation.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum AnimationTrigger {
    /// Fires on the next mouse click (default).
    #[default]
    OnClick,
    /// Starts at the same time as the previous animation (no click needed).
    WithPrevious,
    /// Starts after the previous animation finishes (no click needed).
    AfterPrevious,
}

/// An animation applied to a slide object.
#[derive(Debug, Clone)]
pub struct AnimationEffect {
    /// The specific animation effect to apply.
    pub effect: AnimationEffectType,
    /// Optionally restrict the animation to a text sub-range within the shape.
    /// Only meaningful for text-targeting animations (e.g. FontColor).
    pub text_target: Option<TextTarget>,
    /// When `Some(n)`, all objects with the same group number on the same slide
    /// appear together on a single click rather than requiring one click each.
    /// Objects with `None` each require their own click (original behaviour).
    pub click_group: Option<u32>,
    /// When this animation fires relative to the previous one.
    pub trigger: AnimationTrigger,
    /// Delay in milliseconds before the animation starts (after its trigger fires).
    pub delay_ms: Option<u32>,
}

/// All supported animation effect types.
#[derive(Debug, Clone, PartialEq)]
pub enum AnimationEffectType {
    // ── Entrance (Basic) ──────────────────────────────────────
    /// Instantly appears (no duration, object hidden until clicked).
    Appear,
    /// Fades in from transparent.
    FadeIn,
    /// Flies in from the specified edge.
    FlyIn(Direction),
    /// Wipes in from the specified edge.
    WipeIn(Direction),
    /// Zooms in from a small size.
    ZoomIn,
    /// Barn-door split entrance (two halves move apart).
    SplitIn(SplitOrientation),
    /// Horizontal or vertical blinds entrance.
    BlindsIn(SplitOrientation),
    /// Checkerboard pattern entrance.
    CheckerboardIn(CheckerboardDir),
    /// Pixels dissolve in randomly.
    DissolveIn,
    /// Peeks in from the specified edge.
    PeekIn(Direction),
    /// Random horizontal or vertical bars appear.
    RandomBarsIn(SplitOrientation),
    /// Shape (Box/Circle/Diamond/Plus) entrance.
    ShapeIn(ShapeVariant),
    /// Diagonal strips entrance.
    StripsIn(StripDir),
    /// Wedge (pie-slice) entrance.
    WedgeIn,
    /// Wheel/pinwheel entrance with N spokes (1, 2, 3, 4, or 8).
    WheelIn(u32),

    // ── Entrance (Subtle) ─────────────────────────────────────
    /// Object expands from 0 to full size.
    ExpandIn,
    /// Object swivels (flips around vertical axis) into view.
    SwivelIn,
    /// Zooms in from a small size (Basic Zoom variant).
    BasicZoomIn,

    // ── Entrance (Moderate) ───────────────────────────────────
    /// Revolves in while growing (Centre Revolve).
    CentreRevolveIn,
    /// Floats in from the specified edge with a simultaneous fade.
    FloatIn(Direction),
    /// Grows from small while turning (scale + rotation).
    GrowTurnIn,
    /// Rises up from below.
    RiseUpIn,
    /// Scales in while spinning (Spinner).
    SpinnerIn,
    /// Stretches in along one axis from the specified direction.
    StretchIn(Direction),

    // ── Entrance (Exciting) ───────────────────────────────────
    /// Boomerang curved-path entrance.
    BoomerangIn,
    /// Object bounces in from above.
    BounceIn,
    /// Credits-style upward-scroll entrance.
    CreditsIn,
    /// Curves up from below.
    CurveUpIn,
    /// Drops in from the top.
    DropIn,
    /// Flips into view (approximated as 3D-style rotation).
    FlipIn,
    /// Fast pinwheel (rotation + scale).
    PinwheelIn,
    /// Spirals in to final position.
    SpiralIn,
    /// Basic swivel rotation entrance.
    BasicSwivelIn,
    /// Whip entrance (fast curved motion from right).
    WhipIn,

    // ── Exit (Basic) ─────────────────────────────────────────
    /// Instantly disappears (no duration).
    Disappear,
    /// Fades out to transparent.
    FadeOut,
    /// Flies out toward the specified edge.
    FlyOut(Direction),
    /// Wipes out toward the specified edge.
    WipeOut(Direction),
    /// Zooms out to a small size.
    ZoomOut,
    /// Barn-door split exit (two halves move together).
    SplitOut(SplitOrientation),
    /// Horizontal or vertical blinds exit.
    BlindsOut(SplitOrientation),
    /// Checkerboard pattern exit.
    CheckerboardOut(CheckerboardDir),
    /// Pixels dissolve out randomly.
    DissolveOut,
    /// Peeks out toward the specified edge.
    PeekOut(Direction),
    /// Random horizontal or vertical bars disappear.
    RandomBarsOut(SplitOrientation),
    /// Shape (Box/Circle/Diamond/Plus) exit.
    ShapeOut(ShapeVariant),
    /// Diagonal strips exit.
    StripsOut(StripDir),
    /// Wedge (pie-slice) exit.
    WedgeOut,
    /// Wheel/pinwheel exit with N spokes (1, 2, 3, 4, or 8).
    WheelOut(u32),

    // ── Exit (Subtle) ─────────────────────────────────────────
    /// Contracts from full size to nothing (exit of Expand).
    ContractOut,
    /// Swivels (flips around vertical axis) out of view.
    SwivelOut,

    // ── Exit (Moderate) ───────────────────────────────────────
    /// Revolves out while shrinking (Centre Revolve exit).
    CentreRevolveOut,
    /// Collapses vertically — height shrinks to zero.
    CollapseOut,
    /// Floats out toward the specified edge with a simultaneous fade.
    FloatOut(Direction),
    /// Shrinks while turning (scale + rotation exit).
    ShrinkTurnOut,
    /// Sinks down off the bottom of the slide.
    SinkDownOut,
    /// Scales out while spinning (Spinner exit).
    SpinnerOut,
    /// Zooms out to a small size (Basic Zoom exit variant).
    BasicZoomOut,
    /// Stretches out along one axis toward the specified direction.
    StretchyOut(Direction),

    // ── Exit (Exciting) ───────────────────────────────────────
    /// Boomerang curved-path exit.
    BoomerangOut,
    /// Object bounces up once, then exits downward.
    BounceOut,
    /// Credits-style upward-scroll exit (continues off-screen).
    CreditsOut,
    /// Curves down and off-screen (exit mirror of Curve Up).
    CurveDownOut,
    /// Drops off the bottom of the slide.
    DropOut,
    /// Flips out of view (horizontal swivel to zero width).
    FlipOut,
    /// Fast pinwheel exit (rotation + scale down).
    PinwheelOut,
    /// Spirals out from final position.
    SpiralOut,
    /// Basic swivel rotation exit.
    BasicSwivelOut,
    /// Whip exit (fast curved motion to the right).
    WhipOut,

    // ── Emphasis (Basic) ──────────────────────────────────────
    /// Spin by `degrees` clockwise (e.g. `360.0` for one full rotation).
    Spin(f32),
    /// Brief grow-then-shrink pulse (remains at original size).
    Pulse,
    /// Grow or shrink to `scale` factor (e.g. `1.5` = 150%).
    GrowShrink(f32),
    /// Changes the fill colour to the given hex value (permanent hold).
    FillColor(String),
    /// Changes the font/text colour to the given hex value (permanent hold).
    FontColor(String),
    /// Changes the line/stroke colour to the given hex value (permanent hold).
    LineColor(String),
    /// Pulses opacity from 1.0 down to `level` (0.0–1.0) and back.
    Transparency(f32),

    // ── Emphasis (Subtle) ─────────────────────────────────────
    /// Briefly makes the text bold and reverts.
    BoldFlash,
    /// Brush-style fill-colour flash: changes fill to hex and auto-reverses.
    BrushColor(String),
    /// Shifts the fill colour by 180° in hue (complementary colour) and reverts.
    ComplementaryColor,
    /// Shifts the fill colour by 120° in hue (second complementary) and reverts.
    ComplementaryColor2,
    /// Flashes the fill colour toward a contrasting luminance and reverts.
    ContrastingColor,
    /// Temporarily darkens the fill (opacity drop to 55 % and back).
    Darken,
    /// Temporarily desaturates the fill colour toward grey and reverts.
    Desaturate,
    /// Temporarily lightens the fill colour and reverts.
    Lighten,
    /// Changes the fill colour to the given hex value and auto-reverses (object-colour style).
    ObjectColor(String),
    /// Briefly underlines the text and reverts.
    Underline,

    // ── Emphasis (Moderate) ───────────────────────────────────
    /// Pulses the fill colour to the given hex colour and back.
    ColorPulse(String),
    /// Scales to 115 % while changing fill to the given hex; both stay (hold).
    GrowWithColor(String),
    /// Rapid multi-cycle opacity shimmer.
    Shimmer,
    /// Small left-right teetering rotation (4° × 4 swings).
    Teeter,

    // ── Emphasis (Exciting) ───────────────────────────────────
    /// Rapidly blinks visibility on/off (3 blink cycles).
    Blink,
    /// Bold flash combined with a brief scale-up.
    BoldReveal,
    /// Wave-like oscillating rotation (5 swings, decreasing amplitude).
    Wave,
}

impl AnimationEffect {
    fn new(effect: AnimationEffectType) -> Self {
        Self { effect, text_target: None, click_group: None, trigger: AnimationTrigger::OnClick, delay_ms: None }
    }

    /// Group this animation with others sharing the same `group` number so they
    /// all appear on a single click rather than one click each.
    pub fn with_group(mut self, group: u32) -> Self {
        self.click_group = Some(group);
        self
    }

    /// Start this animation at the same time as the previous one (no click needed).
    pub fn with_previous(mut self) -> Self {
        self.trigger = AnimationTrigger::WithPrevious;
        self
    }

    /// Start this animation after the previous one finishes (no click needed).
    pub fn after_previous(mut self) -> Self {
        self.trigger = AnimationTrigger::AfterPrevious;
        self
    }

    /// Add a delay (in milliseconds) before the animation starts after its trigger fires.
    pub fn delay(mut self, ms: u32) -> Self {
        self.delay_ms = Some(ms);
        self
    }

    /// Restrict the animation to characters `st_idx..=end_idx` within the shape's text.
    pub fn with_char_range(mut self, st_idx: u32, end_idx: u32) -> Self {
        self.text_target = Some(TextTarget::CharRange { st_idx, end_idx });
        self
    }

    /// Restrict the animation to paragraphs `st_idx..=end_idx` within the shape's text.
    pub fn with_para_range(mut self, st_idx: u32, end_idx: u32) -> Self {
        self.text_target = Some(TextTarget::ParaRange { st_idx, end_idx });
        self
    }

    // ── Entrance (basic) ─────────────────────────────────────
    /// Create an Appear entrance animation.
    pub fn appear()    -> Self { Self::new(AnimationEffectType::Appear) }
    /// Create a Fade-In entrance animation.
    pub fn fade_in()   -> Self { Self::new(AnimationEffectType::FadeIn) }
    /// Create a Fly-In entrance animation from the given direction.
    pub fn fly_in(dir: Direction)  -> Self { Self::new(AnimationEffectType::FlyIn(dir)) }
    /// Create a Wipe-In entrance animation from the given direction.
    pub fn wipe_in(dir: Direction) -> Self { Self::new(AnimationEffectType::WipeIn(dir)) }
    /// Create a Zoom-In entrance animation.
    pub fn zoom_in()   -> Self { Self::new(AnimationEffectType::ZoomIn) }
    /// Create a Split (barn-door) entrance animation.
    pub fn split_in(o: SplitOrientation) -> Self { Self::new(AnimationEffectType::SplitIn(o)) }
    /// Create a Blinds entrance animation.
    pub fn blinds_in(o: SplitOrientation)    -> Self { Self::new(AnimationEffectType::BlindsIn(o)) }
    /// Create a Checkerboard entrance animation.
    pub fn checkerboard_in(d: CheckerboardDir) -> Self { Self::new(AnimationEffectType::CheckerboardIn(d)) }
    /// Create a Dissolve-In entrance animation.
    pub fn dissolve_in()                     -> Self { Self::new(AnimationEffectType::DissolveIn) }
    /// Create a Peek-In entrance animation from the given direction.
    pub fn peek_in(dir: Direction)           -> Self { Self::new(AnimationEffectType::PeekIn(dir)) }
    /// Create a Random-Bars entrance animation.
    pub fn random_bars_in(o: SplitOrientation) -> Self { Self::new(AnimationEffectType::RandomBarsIn(o)) }
    /// Create a Shape entrance animation with the given variant.
    pub fn shape_in(v: ShapeVariant)         -> Self { Self::new(AnimationEffectType::ShapeIn(v)) }
    /// Create a Strips entrance animation in the given diagonal direction.
    pub fn strips_in(d: StripDir)            -> Self { Self::new(AnimationEffectType::StripsIn(d)) }
    /// Create a Wedge entrance animation.
    pub fn wedge_in()                        -> Self { Self::new(AnimationEffectType::WedgeIn) }
    /// Create a Wheel entrance animation with the given number of spokes.
    pub fn wheel_in(spokes: u32)             -> Self { Self::new(AnimationEffectType::WheelIn(spokes)) }

    // ── Entrance (subtle) ────────────────────────────────────
    /// Create an Expand-In entrance animation.
    pub fn expand_in()    -> Self { Self::new(AnimationEffectType::ExpandIn) }
    /// Create a Swivel-In entrance animation.
    pub fn swivel_in()    -> Self { Self::new(AnimationEffectType::SwivelIn) }
    /// Create a Basic-Zoom-In entrance animation.
    pub fn basic_zoom_in() -> Self { Self::new(AnimationEffectType::BasicZoomIn) }

    // ── Entrance (moderate) ──────────────────────────────────
    /// Create a Centre-Revolve entrance animation.
    pub fn centre_revolve_in()      -> Self { Self::new(AnimationEffectType::CentreRevolveIn) }
    /// Create a Float-In entrance animation from the given direction.
    pub fn float_in(dir: Direction) -> Self { Self::new(AnimationEffectType::FloatIn(dir)) }
    /// Create a Grow-and-Turn entrance animation.
    pub fn grow_turn_in()           -> Self { Self::new(AnimationEffectType::GrowTurnIn) }
    /// Create a Rise-Up entrance animation.
    pub fn rise_up_in()             -> Self { Self::new(AnimationEffectType::RiseUpIn) }
    /// Create a Spinner entrance animation.
    pub fn spinner_in()             -> Self { Self::new(AnimationEffectType::SpinnerIn) }
    /// Create a Stretch-In entrance animation from the given direction.
    pub fn stretch_in(dir: Direction) -> Self { Self::new(AnimationEffectType::StretchIn(dir)) }

    // ── Entrance (exciting) ──────────────────────────────────
    /// Create a Boomerang entrance animation.
    pub fn boomerang_in()   -> Self { Self::new(AnimationEffectType::BoomerangIn) }
    /// Create a Bounce-In entrance animation.
    pub fn bounce_in()      -> Self { Self::new(AnimationEffectType::BounceIn) }
    /// Create a Credits entrance animation.
    pub fn credits_in()     -> Self { Self::new(AnimationEffectType::CreditsIn) }
    /// Create a Curve-Up entrance animation.
    pub fn curve_up_in()    -> Self { Self::new(AnimationEffectType::CurveUpIn) }
    /// Create a Drop-In entrance animation.
    pub fn drop_in()        -> Self { Self::new(AnimationEffectType::DropIn) }
    /// Create a Flip entrance animation.
    pub fn flip_in()        -> Self { Self::new(AnimationEffectType::FlipIn) }
    /// Create a Pinwheel entrance animation.
    pub fn pinwheel_in()    -> Self { Self::new(AnimationEffectType::PinwheelIn) }
    /// Create a Spiral entrance animation.
    pub fn spiral_in()      -> Self { Self::new(AnimationEffectType::SpiralIn) }
    /// Create a Basic-Swivel entrance animation.
    pub fn basic_swivel_in() -> Self { Self::new(AnimationEffectType::BasicSwivelIn) }
    /// Create a Whip entrance animation.
    pub fn whip_in()        -> Self { Self::new(AnimationEffectType::WhipIn) }

    // ── Exit (basic) ─────────────────────────────────────────
    /// Create a Disappear exit animation.
    pub fn disappear() -> Self { Self::new(AnimationEffectType::Disappear) }
    /// Create a Fade-Out exit animation.
    pub fn fade_out()  -> Self { Self::new(AnimationEffectType::FadeOut) }
    /// Create a Fly-Out exit animation toward the given direction.
    pub fn fly_out(dir: Direction)  -> Self { Self::new(AnimationEffectType::FlyOut(dir)) }
    /// Create a Wipe-Out exit animation toward the given direction.
    pub fn wipe_out(dir: Direction) -> Self { Self::new(AnimationEffectType::WipeOut(dir)) }
    /// Create a Zoom-Out exit animation.
    pub fn zoom_out()  -> Self { Self::new(AnimationEffectType::ZoomOut) }
    /// Create a Split (barn-door) exit animation.
    pub fn split_out(o: SplitOrientation) -> Self { Self::new(AnimationEffectType::SplitOut(o)) }
    /// Create a Blinds exit animation.
    pub fn blinds_out(o: SplitOrientation) -> Self { Self::new(AnimationEffectType::BlindsOut(o)) }
    /// Create a Checkerboard exit animation.
    pub fn checkerboard_out(d: CheckerboardDir) -> Self { Self::new(AnimationEffectType::CheckerboardOut(d)) }
    /// Create a Dissolve-Out exit animation.
    pub fn dissolve_out() -> Self { Self::new(AnimationEffectType::DissolveOut) }
    /// Create a Peek-Out exit animation toward the given direction.
    pub fn peek_out(dir: Direction) -> Self { Self::new(AnimationEffectType::PeekOut(dir)) }
    /// Create a Random-Bars exit animation.
    pub fn random_bars_out(o: SplitOrientation) -> Self { Self::new(AnimationEffectType::RandomBarsOut(o)) }
    /// Create a Shape exit animation with the given variant.
    pub fn shape_out(v: ShapeVariant) -> Self { Self::new(AnimationEffectType::ShapeOut(v)) }
    /// Create a Strips exit animation in the given diagonal direction.
    pub fn strips_out(d: StripDir) -> Self { Self::new(AnimationEffectType::StripsOut(d)) }
    /// Create a Wedge exit animation.
    pub fn wedge_out() -> Self { Self::new(AnimationEffectType::WedgeOut) }
    /// Create a Wheel exit animation with the given number of spokes.
    pub fn wheel_out(spokes: u32) -> Self { Self::new(AnimationEffectType::WheelOut(spokes)) }

    // ── Exit (subtle) ────────────────────────────────────────
    /// Create a Contract-Out exit animation.
    pub fn contract_out() -> Self { Self::new(AnimationEffectType::ContractOut) }
    /// Create a Swivel-Out exit animation.
    pub fn swivel_out() -> Self { Self::new(AnimationEffectType::SwivelOut) }

    // ── Exit (moderate) ──────────────────────────────────────
    /// Create a Centre-Revolve exit animation.
    pub fn centre_revolve_out() -> Self { Self::new(AnimationEffectType::CentreRevolveOut) }
    /// Create a Collapse-Out exit animation.
    pub fn collapse_out() -> Self { Self::new(AnimationEffectType::CollapseOut) }
    /// Create a Float-Out exit animation toward the given direction.
    pub fn float_out(dir: Direction) -> Self { Self::new(AnimationEffectType::FloatOut(dir)) }
    /// Create a Shrink-and-Turn exit animation.
    pub fn shrink_turn_out() -> Self { Self::new(AnimationEffectType::ShrinkTurnOut) }
    /// Create a Sink-Down exit animation.
    pub fn sink_down_out() -> Self { Self::new(AnimationEffectType::SinkDownOut) }
    /// Create a Spinner exit animation.
    pub fn spinner_out() -> Self { Self::new(AnimationEffectType::SpinnerOut) }
    /// Create a Basic-Zoom-Out exit animation.
    pub fn basic_zoom_out() -> Self { Self::new(AnimationEffectType::BasicZoomOut) }
    /// Create a Stretch-Out exit animation toward the given direction.
    pub fn stretchy_out(dir: Direction) -> Self { Self::new(AnimationEffectType::StretchyOut(dir)) }

    // ── Exit (exciting) ──────────────────────────────────────
    /// Create a Boomerang exit animation.
    pub fn boomerang_out() -> Self { Self::new(AnimationEffectType::BoomerangOut) }
    /// Create a Bounce-Out exit animation.
    pub fn bounce_out() -> Self { Self::new(AnimationEffectType::BounceOut) }
    /// Create a Credits exit animation.
    pub fn credits_out() -> Self { Self::new(AnimationEffectType::CreditsOut) }
    /// Create a Curve-Down exit animation.
    pub fn curve_down_out() -> Self { Self::new(AnimationEffectType::CurveDownOut) }
    /// Create a Drop-Out exit animation.
    pub fn drop_out() -> Self { Self::new(AnimationEffectType::DropOut) }
    /// Create a Flip exit animation.
    pub fn flip_out() -> Self { Self::new(AnimationEffectType::FlipOut) }
    /// Create a Pinwheel exit animation.
    pub fn pinwheel_out() -> Self { Self::new(AnimationEffectType::PinwheelOut) }
    /// Create a Spiral exit animation.
    pub fn spiral_out() -> Self { Self::new(AnimationEffectType::SpiralOut) }
    /// Create a Basic-Swivel exit animation.
    pub fn basic_swivel_out() -> Self { Self::new(AnimationEffectType::BasicSwivelOut) }
    /// Create a Whip exit animation.
    pub fn whip_out() -> Self { Self::new(AnimationEffectType::WhipOut) }

    // ── Emphasis (basic) ─────────────────────────────────────
    /// Create a Spin emphasis animation by the given degrees (e.g. `360.0` for one full rotation).
    pub fn spin(degrees: f32)       -> Self { Self::new(AnimationEffectType::Spin(degrees)) }
    /// Create a Pulse emphasis animation (brief grow-then-shrink).
    pub fn pulse()                  -> Self { Self::new(AnimationEffectType::Pulse) }
    /// Create a Grow/Shrink emphasis animation to the given scale factor (e.g. `1.5` = 150%).
    pub fn grow_shrink(scale: f32)  -> Self { Self::new(AnimationEffectType::GrowShrink(scale)) }
    /// Create a Fill-Color emphasis animation. Color is hex, no `#` prefix.
    pub fn fill_color(hex: &str)    -> Self { Self::new(AnimationEffectType::FillColor(hex.trim_start_matches('#').to_uppercase())) }
    /// Create a Font-Color emphasis animation. Color is hex, no `#` prefix.
    pub fn font_color(hex: &str)    -> Self { Self::new(AnimationEffectType::FontColor(hex.trim_start_matches('#').to_uppercase())) }
    /// Create a Line-Color emphasis animation. Color is hex, no `#` prefix.
    pub fn line_color(hex: &str)    -> Self { Self::new(AnimationEffectType::LineColor(hex.trim_start_matches('#').to_uppercase())) }
    /// Create a Transparency emphasis animation. Level is 0.0 (opaque) – 1.0 (fully transparent).
    pub fn transparency(level: f32) -> Self { Self::new(AnimationEffectType::Transparency(level)) }

    // ── Emphasis (subtle) ────────────────────────────────────
    /// Create a Bold-Flash emphasis animation.
    pub fn bold_flash()             -> Self { Self::new(AnimationEffectType::BoldFlash) }
    /// Create a Brush-Color emphasis animation. Color is hex, no `#` prefix.
    pub fn brush_color(hex: &str)   -> Self { Self::new(AnimationEffectType::BrushColor(hex.trim_start_matches('#').to_uppercase())) }
    /// Create a Complementary-Color emphasis animation.
    pub fn complementary_color()    -> Self { Self::new(AnimationEffectType::ComplementaryColor) }
    /// Create a second Complementary-Color emphasis animation (120° hue shift).
    pub fn complementary_color2()   -> Self { Self::new(AnimationEffectType::ComplementaryColor2) }
    /// Create a Contrasting-Color emphasis animation.
    pub fn contrasting_color()      -> Self { Self::new(AnimationEffectType::ContrastingColor) }
    /// Create a Darken emphasis animation.
    pub fn darken()                 -> Self { Self::new(AnimationEffectType::Darken) }
    /// Create a Desaturate emphasis animation.
    pub fn desaturate()             -> Self { Self::new(AnimationEffectType::Desaturate) }
    /// Create a Lighten emphasis animation.
    pub fn lighten()                -> Self { Self::new(AnimationEffectType::Lighten) }
    /// Create an Object-Color emphasis animation. Color is hex, no `#` prefix.
    pub fn object_color(hex: &str)  -> Self { Self::new(AnimationEffectType::ObjectColor(hex.trim_start_matches('#').to_uppercase())) }
    /// Create an Underline emphasis animation.
    pub fn underline()              -> Self { Self::new(AnimationEffectType::Underline) }

    // ── Emphasis (moderate) ──────────────────────────────────
    /// Create a Color-Pulse emphasis animation. Color is hex, no `#` prefix.
    pub fn color_pulse(hex: &str)   -> Self { Self::new(AnimationEffectType::ColorPulse(hex.trim_start_matches('#').to_uppercase())) }
    /// Create a Grow-with-Color emphasis animation. Color is hex, no `#` prefix.
    pub fn grow_with_color(hex: &str) -> Self { Self::new(AnimationEffectType::GrowWithColor(hex.trim_start_matches('#').to_uppercase())) }
    /// Create a Shimmer emphasis animation.
    pub fn shimmer()                -> Self { Self::new(AnimationEffectType::Shimmer) }
    /// Create a Teeter emphasis animation.
    pub fn teeter()                 -> Self { Self::new(AnimationEffectType::Teeter) }

    // ── Emphasis (exciting) ──────────────────────────────────
    /// Create a Blink emphasis animation.
    pub fn blink()                  -> Self { Self::new(AnimationEffectType::Blink) }
    /// Create a Bold-Reveal emphasis animation.
    pub fn bold_reveal()            -> Self { Self::new(AnimationEffectType::BoldReveal) }
    /// Create a Wave emphasis animation.
    pub fn wave()                   -> Self { Self::new(AnimationEffectType::Wave) }
}

// ─────────────────────────────────────────────────────────────
// Slide transitions
// ─────────────────────────────────────────────────────────────

/// Slide transition type.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum TransitionType {
    /// No transition element emitted.
    /// No transition element emitted.
    #[default]
    None,
    /// Cut transition (instant switch).
    Cut,
    /// Fade transition.
    Fade,
    /// Push transition.
    Push,
    /// Wipe transition.
    Wipe,
    /// Split (barn-door) transition.
    Split,
    /// Cover transition.
    Cover,
    /// Uncover transition.
    Uncover,
    /// Zoom transition.
    Zoom,
    /// Flash transition.
    Flash,
    /// Morph transition.
    Morph,
    /// Vortex transition.
    Vortex,
    /// Ripple transition.
    Ripple,
    /// Glitter transition.
    Glitter,
    /// Honeycomb transition.
    Honeycomb,
    /// Shred transition.
    Shred,
    /// Switch transition.
    Switch,
    /// Flip transition.
    Flip,
    /// Pan transition.
    Pan,
    /// Ferris wheel transition.
    Ferris,
    /// Gallery transition.
    Gallery,
    /// Conveyor transition.
    Conveyor,
    /// Doors transition.
    Doors,
    /// Box transition.
    Box,
    /// Random transition.
    Random,
    /// Random bars transition.
    RandomBar,
    /// Circle transition.
    Circle,
    /// Diamond transition.
    Diamond,
    /// Wheel transition.
    Wheel,
    /// Checker transition.
    Checker,
    /// Blinds transition.
    Blinds,
    /// Strips transition.
    Strips,
    /// Plus transition.
    Plus,
}

/// Direction hint for transitions that support it.
#[derive(Debug, Clone, PartialEq)]
pub enum TransitionDir {
    /// Left direction.
    Left,
    /// Right direction.
    Right,
    /// Up direction.
    Up,
    /// Down direction.
    Down,
    /// Left-down diagonal direction.
    LeftDown,
    /// Left-up diagonal direction.
    LeftUp,
    /// Right-down diagonal direction.
    RightDown,
    /// Right-up diagonal direction.
    RightUp,
}

impl TransitionDir {
    /// Return the OOXML direction string (e.g. `"l"`, `"r"`, `"u"`, `"d"`).
    pub fn as_str(&self) -> &'static str {
        match self {
            TransitionDir::Left      => "l",
            TransitionDir::Right     => "r",
            TransitionDir::Up        => "u",
            TransitionDir::Down      => "d",
            TransitionDir::LeftDown  => "ld",
            TransitionDir::LeftUp    => "lu",
            TransitionDir::RightDown => "rd",
            TransitionDir::RightUp   => "ru",
        }
    }
}

/// Speed preset for a transition.
#[derive(Debug, Clone, PartialEq)]
pub enum TransitionSpeed {
    /// Slow transition speed.
    Slow,
    /// Medium transition speed.
    Medium,
    /// Fast transition speed.
    Fast,
}

impl TransitionSpeed {
    /// Return the OOXML speed string (`"slow"`, `"med"`, or `"fast"`).
    pub fn as_str(&self) -> &'static str {
        match self {
            TransitionSpeed::Slow   => "slow",
            TransitionSpeed::Medium => "med",
            TransitionSpeed::Fast   => "fast",
        }
    }
}

/// Slide transition properties.
#[derive(Debug, Clone, Default)]
pub struct TransitionProps {
    /// The transition effect type.
    pub transition_type: TransitionType,
    /// Optional direction hint (applies to Push, Wipe, Cover, etc.)
    pub direction: Option<TransitionDir>,
    /// Speed preset (slow/med/fast). Ignored when `duration_ms` is set.
    pub speed: Option<TransitionSpeed>,
    /// Explicit transition duration in milliseconds.
    pub duration_ms: Option<u32>,
    /// Auto-advance after this many milliseconds (in addition to or instead of click).
    pub advance_after_ms: Option<u32>,
    /// Advance on mouse click (default: true).
    pub advance_on_click: bool,
}

impl TransitionProps {
    /// Create transition properties with the given type and default settings.
    pub fn new(t: TransitionType) -> Self {
        TransitionProps { transition_type: t, advance_on_click: true, ..Default::default() }
    }
    /// Create a fade transition with default settings.
    pub fn fade() -> Self { Self::new(TransitionType::Fade) }
    /// Create a push transition in the given direction.
    pub fn push(dir: TransitionDir) -> Self {
        let mut p = Self::new(TransitionType::Push);
        p.direction = Some(dir);
        p
    }
    /// Create a wipe transition in the given direction.
    pub fn wipe(dir: TransitionDir) -> Self {
        let mut p = Self::new(TransitionType::Wipe);
        p.direction = Some(dir);
        p
    }
}
