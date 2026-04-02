/// One inch in English Metric Units (EMU)
pub const EMU: i64 = 914_400;
/// One point (pt) in EMU
pub const ONEPT: i64 = 12_700;
/// Golden ratio modifier for line-height
pub const LINEH_MODIFIER: f64 = 1.67;
/// Default bullet indent
pub const DEF_BULLET_MARGIN: i64 = 27;
/// Default font color (hex, no #)
pub const DEF_FONT_COLOR: &str = "000000";
/// Default font size (points)
pub const DEF_FONT_SIZE: f64 = 12.0;
/// Default title font size (points)
pub const DEF_FONT_TITLE_SIZE: f64 = 18.0;
/// Default shape line color
pub const DEF_SHAPE_LINE_COLOR: &str = "333333";
/// Default slide background color
pub const DEF_SLIDE_BKGD: &str = "FFFFFF";
/// Default slide margins (TRBL, inches)
pub const DEF_SLIDE_MARGIN_IN: [f64; 4] = [0.5, 0.5, 0.5, 0.5];
/// Default cell margins (TRBL, inches)
pub const DEF_CELL_MARGIN_IN: [f64; 4] = [0.05, 0.1, 0.05, 0.1];
/// Slide layout series base ID
pub const LAYOUT_IDX_SERIES_BASE: u64 = 2_147_483_649;
/// Slide number field ID
pub const SLDNUMFLDID: &str = "{F7021451-1387-4CA6-816F-3879F97B5CBC}";
/// Hex color regex pattern (6 hex chars)
pub const REGEX_HEX_PATTERN: &str = "^[0-9a-fA-F]{6}$";

/// Validates that a string is a 6-digit hex color
pub fn is_hex_color(s: &str) -> bool {
    s.len() == 6 && s.chars().all(|c| c.is_ascii_hexdigit())
}

/// Scheme colors (theme colors) as used in OOXML
#[derive(Debug, Clone, PartialEq)]
pub enum SchemeColor {
    /// text1 / dark text
    Text1,
    /// text2
    Text2,
    /// background1
    Background1,
    /// background2
    Background2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
}

impl SchemeColor {
    /// Returns the OOXML string value (e.g. "tx1", "accent1")
    pub fn as_str(&self) -> &'static str {
        match self {
            SchemeColor::Text1 => "tx1",
            SchemeColor::Text2 => "tx2",
            SchemeColor::Background1 => "bg1",
            SchemeColor::Background2 => "bg2",
            SchemeColor::Accent1 => "accent1",
            SchemeColor::Accent2 => "accent2",
            SchemeColor::Accent3 => "accent3",
            SchemeColor::Accent4 => "accent4",
            SchemeColor::Accent5 => "accent5",
            SchemeColor::Accent6 => "accent6",
        }
    }

    /// Try to parse from an OOXML scheme color string
    pub fn from_str(s: &str) -> Option<Self> {
        match s {
            "tx1" => Some(SchemeColor::Text1),
            "tx2" => Some(SchemeColor::Text2),
            "bg1" => Some(SchemeColor::Background1),
            "bg2" => Some(SchemeColor::Background2),
            "accent1" => Some(SchemeColor::Accent1),
            "accent2" => Some(SchemeColor::Accent2),
            "accent3" => Some(SchemeColor::Accent3),
            "accent4" => Some(SchemeColor::Accent4),
            "accent5" => Some(SchemeColor::Accent5),
            "accent6" => Some(SchemeColor::Accent6),
            _ => None,
        }
    }
}

/// Horizontal text alignment
#[derive(Debug, Clone, PartialEq, Default)]
pub enum AlignH {
    #[default]
    Left,
    Center,
    Right,
    Justify,
    Distribute,
}

impl AlignH {
    pub fn as_ooxml(&self) -> &'static str {
        match self {
            AlignH::Left => "l",
            AlignH::Center => "ctr",
            AlignH::Right => "r",
            AlignH::Justify => "just",
            AlignH::Distribute => "dist",
        }
    }
}

/// Vertical text alignment
#[derive(Debug, Clone, PartialEq, Default)]
pub enum AlignV {
    Top,
    #[default]
    Middle,
    Bottom,
}

impl AlignV {
    pub fn as_ooxml(&self) -> &'static str {
        match self {
            AlignV::Top => "t",
            AlignV::Middle => "ctr",
            AlignV::Bottom => "b",
        }
    }
}

/// Chart types
#[derive(Debug, Clone, PartialEq)]
pub enum ChartType {
    Area,
    Bar,
    Bar3D,
    Bubble,
    Bubble3D,
    Doughnut,
    Line,
    Pie,
    Radar,
    Scatter,
    /// Stock chart — High, Low, Close (3 series)
    StockHLC,
    /// Stock chart — Open, High, Low, Close (4 series)
    StockOHLC,
    /// Stock chart — Volume + High, Low, Close (4 series, volume bar)
    StockVHLC,
    /// Stock chart — Volume + Open, High, Low, Close (5 series, volume bar)
    StockVOHLC,
    /// 3D surface chart
    Surface,
    /// 3D surface wireframe chart
    SurfaceWireframe,
    /// 2D top-view surface chart
    SurfaceTop,
    /// 2D top-view wireframe surface chart
    SurfaceTopWireframe,
}

/// All 150+ OOXML preset shape types
#[derive(Debug, Clone, PartialEq)]
pub enum ShapeType {
    AccentBorderCallout1,
    AccentBorderCallout2,
    AccentBorderCallout3,
    AccentCallout1,
    AccentCallout2,
    AccentCallout3,
    ActionButtonBackPrevious,
    ActionButtonBeginning,
    ActionButtonBlank,
    ActionButtonDocument,
    ActionButtonEnd,
    ActionButtonForwardNext,
    ActionButtonHelp,
    ActionButtonHome,
    ActionButtonInformation,
    ActionButtonMovie,
    ActionButtonReturn,
    ActionButtonSound,
    Arc,
    BentArrow,
    BentUpArrow,
    Bevel,
    BlockArc,
    BorderCallout1,
    BorderCallout2,
    BorderCallout3,
    BracePair,
    BracketPair,
    Callout1,
    Callout2,
    Callout3,
    Can,
    ChartPlus,
    ChartStar,
    ChartX,
    Chevron,
    Chord,
    CircularArrow,
    Cloud,
    CloudCallout,
    Corner,
    CornerTabs,
    Cube,
    CurvedDownArrow,
    CurvedLeftArrow,
    CurvedRightArrow,
    CurvedUpArrow,
    CustGeom,
    Decagon,
    DiagStripe,
    Diamond,
    Dodecagon,
    Donut,
    DoubleWave,
    DownArrow,
    DownArrowCallout,
    Ellipse,
    EllipseRibbon,
    EllipseRibbon2,
    FlowChartAlternateProcess,
    FlowChartCollate,
    FlowChartConnector,
    FlowChartDecision,
    FlowChartDelay,
    FlowChartDisplay,
    FlowChartDocument,
    FlowChartExtract,
    FlowChartInputOutput,
    FlowChartInternalStorage,
    FlowChartMagneticDisk,
    FlowChartMagneticDrum,
    FlowChartMagneticTape,
    FlowChartManualInput,
    FlowChartManualOperation,
    FlowChartMerge,
    FlowChartMultidocument,
    FlowChartOfflineStorage,
    FlowChartOffpageConnector,
    FlowChartOnlineStorage,
    FlowChartOr,
    FlowChartPredefinedProcess,
    FlowChartPreparation,
    FlowChartProcess,
    FlowChartPunchedCard,
    FlowChartPunchedTape,
    FlowChartSort,
    FlowChartSummingJunction,
    FlowChartTerminator,
    FolderCorner,
    Frame,
    Funnel,
    Gear6,
    Gear9,
    HalfFrame,
    Heart,
    Heptagon,
    Hexagon,
    HomePlate,
    HorizontalScroll,
    IrregularSeal1,
    IrregularSeal2,
    LeftArrow,
    LeftArrowCallout,
    LeftBrace,
    LeftBracket,
    LeftCircularArrow,
    LeftRightArrow,
    LeftRightArrowCallout,
    LeftRightCircularArrow,
    LeftRightRibbon,
    LeftRightUpArrow,
    LeftUpArrow,
    LightningBolt,
    Line,
    LineInv,
    MathDivide,
    MathEqual,
    MathMinus,
    MathMultiply,
    MathNotEqual,
    MathPlus,
    Moon,
    NoSmoking,
    NonIsoscelesTrapezoid,
    NotchedRightArrow,
    Octagon,
    Parallelogram,
    Pentagon,
    Pie,
    PieWedge,
    Plaque,
    PlaqueTabs,
    Plus,
    QuadArrow,
    QuadArrowCallout,
    Rect,
    Ribbon,
    Ribbon2,
    RightArrow,
    RightArrowCallout,
    RightBrace,
    RightBracket,
    Round1Rect,
    Round2DiagRect,
    Round2SameRect,
    RoundRect,
    RtTriangle,
    SmileyFace,
    Snip1Rect,
    Snip2DiagRect,
    Snip2SameRect,
    SnipRoundRect,
    SquareTabs,
    Star10,
    Star12,
    Star16,
    Star24,
    Star32,
    Star4,
    Star5,
    Star6,
    Star7,
    Star8,
    StripedRightArrow,
    Sun,
    SwooshArrow,
    Teardrop,
    Trapezoid,
    Triangle,
    UpArrow,
    UpArrowCallout,
    UpDownArrow,
    UpDownArrowCallout,
    UturnArrow,
    VerticalScroll,
    Wave,
    WedgeEllipseCallout,
    WedgeRectCallout,
    WedgeRoundRectCallout,
}

impl ShapeType {
    /// Returns the OOXML preset geometry name string
    pub fn as_str(&self) -> &'static str {
        match self {
            ShapeType::AccentBorderCallout1 => "accentBorderCallout1",
            ShapeType::AccentBorderCallout2 => "accentBorderCallout2",
            ShapeType::AccentBorderCallout3 => "accentBorderCallout3",
            ShapeType::AccentCallout1 => "accentCallout1",
            ShapeType::AccentCallout2 => "accentCallout2",
            ShapeType::AccentCallout3 => "accentCallout3",
            ShapeType::ActionButtonBackPrevious => "actionButtonBackPrevious",
            ShapeType::ActionButtonBeginning => "actionButtonBeginning",
            ShapeType::ActionButtonBlank => "actionButtonBlank",
            ShapeType::ActionButtonDocument => "actionButtonDocument",
            ShapeType::ActionButtonEnd => "actionButtonEnd",
            ShapeType::ActionButtonForwardNext => "actionButtonForwardNext",
            ShapeType::ActionButtonHelp => "actionButtonHelp",
            ShapeType::ActionButtonHome => "actionButtonHome",
            ShapeType::ActionButtonInformation => "actionButtonInformation",
            ShapeType::ActionButtonMovie => "actionButtonMovie",
            ShapeType::ActionButtonReturn => "actionButtonReturn",
            ShapeType::ActionButtonSound => "actionButtonSound",
            ShapeType::Arc => "arc",
            ShapeType::BentArrow => "bentArrow",
            ShapeType::BentUpArrow => "bentUpArrow",
            ShapeType::Bevel => "bevel",
            ShapeType::BlockArc => "blockArc",
            ShapeType::BorderCallout1 => "borderCallout1",
            ShapeType::BorderCallout2 => "borderCallout2",
            ShapeType::BorderCallout3 => "borderCallout3",
            ShapeType::BracePair => "bracePair",
            ShapeType::BracketPair => "bracketPair",
            ShapeType::Callout1 => "callout1",
            ShapeType::Callout2 => "callout2",
            ShapeType::Callout3 => "callout3",
            ShapeType::Can => "can",
            ShapeType::ChartPlus => "chartPlus",
            ShapeType::ChartStar => "chartStar",
            ShapeType::ChartX => "chartX",
            ShapeType::Chevron => "chevron",
            ShapeType::Chord => "chord",
            ShapeType::CircularArrow => "circularArrow",
            ShapeType::Cloud => "cloud",
            ShapeType::CloudCallout => "cloudCallout",
            ShapeType::Corner => "corner",
            ShapeType::CornerTabs => "cornerTabs",
            ShapeType::Cube => "cube",
            ShapeType::CurvedDownArrow => "curvedDownArrow",
            ShapeType::CurvedLeftArrow => "curvedLeftArrow",
            ShapeType::CurvedRightArrow => "curvedRightArrow",
            ShapeType::CurvedUpArrow => "curvedUpArrow",
            ShapeType::CustGeom => "custGeom",
            ShapeType::Decagon => "decagon",
            ShapeType::DiagStripe => "diagStripe",
            ShapeType::Diamond => "diamond",
            ShapeType::Dodecagon => "dodecagon",
            ShapeType::Donut => "donut",
            ShapeType::DoubleWave => "doubleWave",
            ShapeType::DownArrow => "downArrow",
            ShapeType::DownArrowCallout => "downArrowCallout",
            ShapeType::Ellipse => "ellipse",
            ShapeType::EllipseRibbon => "ellipseRibbon",
            ShapeType::EllipseRibbon2 => "ellipseRibbon2",
            ShapeType::FlowChartAlternateProcess => "flowChartAlternateProcess",
            ShapeType::FlowChartCollate => "flowChartCollate",
            ShapeType::FlowChartConnector => "flowChartConnector",
            ShapeType::FlowChartDecision => "flowChartDecision",
            ShapeType::FlowChartDelay => "flowChartDelay",
            ShapeType::FlowChartDisplay => "flowChartDisplay",
            ShapeType::FlowChartDocument => "flowChartDocument",
            ShapeType::FlowChartExtract => "flowChartExtract",
            ShapeType::FlowChartInputOutput => "flowChartInputOutput",
            ShapeType::FlowChartInternalStorage => "flowChartInternalStorage",
            ShapeType::FlowChartMagneticDisk => "flowChartMagneticDisk",
            ShapeType::FlowChartMagneticDrum => "flowChartMagneticDrum",
            ShapeType::FlowChartMagneticTape => "flowChartMagneticTape",
            ShapeType::FlowChartManualInput => "flowChartManualInput",
            ShapeType::FlowChartManualOperation => "flowChartManualOperation",
            ShapeType::FlowChartMerge => "flowChartMerge",
            ShapeType::FlowChartMultidocument => "flowChartMultidocument",
            ShapeType::FlowChartOfflineStorage => "flowChartOfflineStorage",
            ShapeType::FlowChartOffpageConnector => "flowChartOffpageConnector",
            ShapeType::FlowChartOnlineStorage => "flowChartOnlineStorage",
            ShapeType::FlowChartOr => "flowChartOr",
            ShapeType::FlowChartPredefinedProcess => "flowChartPredefinedProcess",
            ShapeType::FlowChartPreparation => "flowChartPreparation",
            ShapeType::FlowChartProcess => "flowChartProcess",
            ShapeType::FlowChartPunchedCard => "flowChartPunchedCard",
            ShapeType::FlowChartPunchedTape => "flowChartPunchedTape",
            ShapeType::FlowChartSort => "flowChartSort",
            ShapeType::FlowChartSummingJunction => "flowChartSummingJunction",
            ShapeType::FlowChartTerminator => "flowChartTerminator",
            ShapeType::FolderCorner => "folderCorner",
            ShapeType::Frame => "frame",
            ShapeType::Funnel => "funnel",
            ShapeType::Gear6 => "gear6",
            ShapeType::Gear9 => "gear9",
            ShapeType::HalfFrame => "halfFrame",
            ShapeType::Heart => "heart",
            ShapeType::Heptagon => "heptagon",
            ShapeType::Hexagon => "hexagon",
            ShapeType::HomePlate => "homePlate",
            ShapeType::HorizontalScroll => "horizontalScroll",
            ShapeType::IrregularSeal1 => "irregularSeal1",
            ShapeType::IrregularSeal2 => "irregularSeal2",
            ShapeType::LeftArrow => "leftArrow",
            ShapeType::LeftArrowCallout => "leftArrowCallout",
            ShapeType::LeftBrace => "leftBrace",
            ShapeType::LeftBracket => "leftBracket",
            ShapeType::LeftCircularArrow => "leftCircularArrow",
            ShapeType::LeftRightArrow => "leftRightArrow",
            ShapeType::LeftRightArrowCallout => "leftRightArrowCallout",
            ShapeType::LeftRightCircularArrow => "leftRightCircularArrow",
            ShapeType::LeftRightRibbon => "leftRightRibbon",
            ShapeType::LeftRightUpArrow => "leftRightUpArrow",
            ShapeType::LeftUpArrow => "leftUpArrow",
            ShapeType::LightningBolt => "lightningBolt",
            ShapeType::Line => "line",
            ShapeType::LineInv => "lineInv",
            ShapeType::MathDivide => "mathDivide",
            ShapeType::MathEqual => "mathEqual",
            ShapeType::MathMinus => "mathMinus",
            ShapeType::MathMultiply => "mathMultiply",
            ShapeType::MathNotEqual => "mathNotEqual",
            ShapeType::MathPlus => "mathPlus",
            ShapeType::Moon => "moon",
            ShapeType::NoSmoking => "noSmoking",
            ShapeType::NonIsoscelesTrapezoid => "nonIsoscelesTrapezoid",
            ShapeType::NotchedRightArrow => "notchedRightArrow",
            ShapeType::Octagon => "octagon",
            ShapeType::Parallelogram => "parallelogram",
            ShapeType::Pentagon => "pentagon",
            ShapeType::Pie => "pie",
            ShapeType::PieWedge => "pieWedge",
            ShapeType::Plaque => "plaque",
            ShapeType::PlaqueTabs => "plaqueTabs",
            ShapeType::Plus => "plus",
            ShapeType::QuadArrow => "quadArrow",
            ShapeType::QuadArrowCallout => "quadArrowCallout",
            ShapeType::Rect => "rect",
            ShapeType::Ribbon => "ribbon",
            ShapeType::Ribbon2 => "ribbon2",
            ShapeType::RightArrow => "rightArrow",
            ShapeType::RightArrowCallout => "rightArrowCallout",
            ShapeType::RightBrace => "rightBrace",
            ShapeType::RightBracket => "rightBracket",
            ShapeType::Round1Rect => "round1Rect",
            ShapeType::Round2DiagRect => "round2DiagRect",
            ShapeType::Round2SameRect => "round2SameRect",
            ShapeType::RoundRect => "roundRect",
            ShapeType::RtTriangle => "rtTriangle",
            ShapeType::SmileyFace => "smileyFace",
            ShapeType::Snip1Rect => "snip1Rect",
            ShapeType::Snip2DiagRect => "snip2DiagRect",
            ShapeType::Snip2SameRect => "snip2SameRect",
            ShapeType::SnipRoundRect => "snipRoundRect",
            ShapeType::SquareTabs => "squareTabs",
            ShapeType::Star10 => "star10",
            ShapeType::Star12 => "star12",
            ShapeType::Star16 => "star16",
            ShapeType::Star24 => "star24",
            ShapeType::Star32 => "star32",
            ShapeType::Star4 => "star4",
            ShapeType::Star5 => "star5",
            ShapeType::Star6 => "star6",
            ShapeType::Star7 => "star7",
            ShapeType::Star8 => "star8",
            ShapeType::StripedRightArrow => "stripedRightArrow",
            ShapeType::Sun => "sun",
            ShapeType::SwooshArrow => "swooshArrow",
            ShapeType::Teardrop => "teardrop",
            ShapeType::Trapezoid => "trapezoid",
            ShapeType::Triangle => "triangle",
            ShapeType::UpArrow => "upArrow",
            ShapeType::UpArrowCallout => "upArrowCallout",
            ShapeType::UpDownArrow => "upDownArrow",
            ShapeType::UpDownArrowCallout => "upDownArrowCallout",
            ShapeType::UturnArrow => "uturnArrow",
            ShapeType::VerticalScroll => "verticalScroll",
            ShapeType::Wave => "wave",
            ShapeType::WedgeEllipseCallout => "wedgeEllipseCallout",
            ShapeType::WedgeRectCallout => "wedgeRectCallout",
            ShapeType::WedgeRoundRectCallout => "wedgeRoundRectCallout",
        }
    }
}
