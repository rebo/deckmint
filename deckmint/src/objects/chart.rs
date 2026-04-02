use crate::enums::ChartType;
use crate::types::PositionProps;

/// A single data series for a chart
#[derive(Debug, Clone)]
pub struct ChartSeries {
    /// Series name (shown in legend)
    pub name: String,
    /// Category labels (x-axis labels for most charts)
    pub labels: Vec<String>,
    /// Data values
    pub values: Vec<f64>,
    /// Optional override color for this series (6-digit hex, no #)
    pub color: Option<String>,
    /// Bubble sizes (for Bubble/Bubble3D charts only).
    pub sizes: Option<Vec<f64>>,
}

impl ChartSeries {
    /// Create a new data series with a name, category labels, and values.
    pub fn new(name: impl Into<String>, labels: Vec<impl Into<String>>, values: Vec<f64>) -> Self {
        ChartSeries {
            name: name.into(),
            labels: labels.into_iter().map(|l| l.into()).collect(),
            values,
            color: None,
            sizes: None,
        }
    }

    /// Set the override color for this series, 6-digit hex, no `#` prefix.
    pub fn color(mut self, c: impl Into<String>) -> Self {
        self.color = Some(c.into().trim_start_matches('#').to_uppercase());
        self
    }

    /// Set bubble sizes for Bubble/Bubble3D charts.
    pub fn sizes(mut self, sizes: Vec<f64>) -> Self {
        self.sizes = Some(sizes);
        self
    }
}

/// Position of the chart legend relative to the plot area.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum LegendPos {
    /// Below the chart (OOXML "b").
    #[default]
    Bottom,
    /// Above the chart (OOXML "t").
    Top,
    /// Left of the chart (OOXML "l").
    Left,
    /// Right of the chart (OOXML "r").
    Right,
    /// Top-right corner of the chart (OOXML "tr").
    TopRight,
}

impl LegendPos {
    /// Return the OOXML attribute value for this legend position.
    pub fn as_ooxml(&self) -> &'static str {
        match self {
            LegendPos::Bottom => "b",
            LegendPos::Top => "t",
            LegendPos::Left => "l",
            LegendPos::Right => "r",
            LegendPos::TopRight => "tr",
        }
    }
}

/// Bar/column chart orientation.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum BarDir {
    /// Vertical bars (OOXML barDir="col").
    #[default]
    Column,
    /// Horizontal bars (OOXML barDir="bar").
    Bar,
}

/// Grouping mode for bar and column charts.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum BarGrouping {
    /// Side-by-side bars for each category (OOXML "clustered").
    #[default]
    Clustered,
    /// Bars stacked on top of each other (OOXML "stacked").
    Stacked,
    /// Bars stacked and normalized to 100% (OOXML "percentStacked").
    PercentStacked,
}

/// Options for chart placement and styling
#[derive(Debug, Clone)]
pub struct ChartOptions {
    /// Position and dimensions on the slide.
    pub position: PositionProps,
    /// Chart title text.
    pub title: Option<String>,
    /// Whether to display the legend.
    pub show_legend: bool,
    /// Legend placement relative to the plot area.
    pub legend_pos: LegendPos,
    /// Whether to show data value labels on series.
    pub show_value: bool,
    /// Explicit series colors in order, 6-digit hex, no `#` prefix. Wraps if more series than colors.
    pub chart_colors: Vec<String>,
    /// Value axis minimum.
    pub val_axis_min: Option<f64>,
    /// Value axis maximum.
    pub val_axis_max: Option<f64>,
    /// Whether to show horizontal grid lines.
    pub show_grid_lines: bool,
    /// Bar/column direction (Bar/Column charts only).
    pub bar_dir: BarDir,
    /// Bar grouping mode (Bar/Column charts only).
    pub bar_grouping: BarGrouping,
    /// Doughnut hole size, valid range 0--100 (Doughnut charts only; default 50).
    pub hole_size: Option<u32>,
    /// Enable line smoothing (Line charts only).
    pub line_smooth: bool,
    /// Whether to show data point markers (Line charts only).
    pub show_markers: bool,
    /// First slice angle in degrees, valid range 0--360 (Pie/Doughnut; default 0 = 12 o'clock).
    pub first_slice_angle: Option<u32>,
    /// Category axis title.
    pub cat_axis_title: Option<String>,
    /// Value axis title.
    pub val_axis_title: Option<String>,
}

impl Default for ChartOptions {
    fn default() -> Self {
        ChartOptions {
            position: PositionProps::default(),
            title: None,
            show_legend: true,
            legend_pos: LegendPos::default(),
            show_value: false,
            chart_colors: Vec::new(),
            val_axis_min: None,
            val_axis_max: None,
            show_grid_lines: true,
            bar_dir: BarDir::default(),
            bar_grouping: BarGrouping::default(),
            hole_size: None,
            line_smooth: false,
            show_markers: true,
            first_slice_angle: None,
            cat_axis_title: None,
            val_axis_title: None,
        }
    }
}

/// A chart placed on a slide
#[derive(Debug, Clone)]
pub struct ChartObject {
    /// Internal object name for relationship tracking.
    pub object_name: String,
    /// rId for the chart relationship on this slide.
    pub chart_rid: u32,
    /// Type of chart (bar, line, pie, etc.).
    pub chart_type: ChartType,
    /// Data series rendered in this chart.
    pub series: Vec<ChartSeries>,
    /// Placement and styling options for this chart.
    pub options: ChartOptions,
}

/// Fluent builder for chart options
pub struct ChartOptionsBuilder {
    opts: ChartOptions,
}

impl ChartOptionsBuilder {
    /// Create a new builder with default chart options.
    pub fn new() -> Self {
        ChartOptionsBuilder { opts: ChartOptions::default() }
    }

    /// Set the X position in inches.
    pub fn x(mut self, v: f64) -> Self { self.opts.position.x = Some(crate::types::Coord::Inches(v)); self }
    /// Set the Y position in inches.
    pub fn y(mut self, v: f64) -> Self { self.opts.position.y = Some(crate::types::Coord::Inches(v)); self }
    /// Set the width in inches.
    pub fn w(mut self, v: f64) -> Self { self.opts.position.w = Some(crate::types::Coord::Inches(v)); self }
    /// Set the height in inches.
    pub fn h(mut self, v: f64) -> Self { self.opts.position.h = Some(crate::types::Coord::Inches(v)); self }
    /// Set position (x, y) in inches.
    pub fn pos(self, x: f64, y: f64) -> Self {
        self.x(x).y(y)
    }
    /// Set size (width, height) in inches.
    pub fn size(self, w: f64, h: f64) -> Self {
        self.w(w).h(h)
    }

    /// Set the chart title text.
    pub fn title(mut self, t: impl Into<String>) -> Self { self.opts.title = Some(t.into()); self }
    /// Show or hide the legend.
    pub fn show_legend(mut self, v: bool) -> Self { self.opts.show_legend = v; self }
    /// Set the legend position relative to the plot area.
    pub fn legend_pos(mut self, p: LegendPos) -> Self { self.opts.legend_pos = p; self }
    /// Enable data value labels on chart series.
    pub fn show_value(mut self) -> Self { self.opts.show_value = true; self }
    /// Set explicit series colors in order, 6-digit hex, no `#` prefix.
    pub fn chart_colors(mut self, colors: Vec<impl Into<String>>) -> Self {
        self.opts.chart_colors = colors.into_iter().map(|c| c.into().trim_start_matches('#').to_uppercase()).collect();
        self
    }
    /// Set the value axis minimum bound.
    pub fn val_axis_min(mut self, v: f64) -> Self { self.opts.val_axis_min = Some(v); self }
    /// Set the value axis maximum bound.
    pub fn val_axis_max(mut self, v: f64) -> Self { self.opts.val_axis_max = Some(v); self }
    /// Hide horizontal grid lines.
    pub fn no_grid_lines(mut self) -> Self { self.opts.show_grid_lines = false; self }
    /// Set bar/column direction (horizontal or vertical).
    pub fn bar_dir(mut self, d: BarDir) -> Self { self.opts.bar_dir = d; self }
    /// Set bar grouping mode (clustered, stacked, or percent-stacked).
    pub fn bar_grouping(mut self, g: BarGrouping) -> Self { self.opts.bar_grouping = g; self }
    /// Set the doughnut hole size, valid range 0--100.
    pub fn hole_size(mut self, s: u32) -> Self { self.opts.hole_size = Some(s); self }
    /// Enable line smoothing for line charts.
    pub fn line_smooth(mut self) -> Self { self.opts.line_smooth = true; self }
    /// Hide data point markers on line charts.
    pub fn no_markers(mut self) -> Self { self.opts.show_markers = false; self }
    /// Set the first slice angle in degrees for pie/doughnut charts, valid range 0--360.
    pub fn first_slice_angle(mut self, deg: u32) -> Self { self.opts.first_slice_angle = Some(deg); self }
    /// Set the category axis title text.
    pub fn cat_axis_title(mut self, t: impl Into<String>) -> Self { self.opts.cat_axis_title = Some(t.into()); self }
    /// Set the value axis title text.
    pub fn val_axis_title(mut self, t: impl Into<String>) -> Self { self.opts.val_axis_title = Some(t.into()); self }

    /// Consume the builder and return the configured chart options.
    pub fn build(self) -> ChartOptions {
        self.opts
    }
}

impl Default for ChartOptionsBuilder {
    fn default() -> Self { Self::new() }
}

/// Default chart color palette (matching PowerPoint Office theme)
pub static DEFAULT_CHART_COLORS: &[&str] = &[
    "4472C4", "ED7D31", "A9D18E", "FFC000", "5B9BD5", "70AD47",
    "FF0000", "7030A0", "00B0F0", "C55A11", "833C00", "636363",
];
