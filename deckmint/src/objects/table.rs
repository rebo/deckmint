use crate::enums::{AlignH, AlignV};
use crate::types::{AnimationEffect, BorderProps, Coord, GradientFill, HyperlinkProps, Margin, PatternFill, PositionProps, ShadowProps};

/// A table placed on a slide.
#[derive(Debug, Clone)]
pub struct TableObject {
    /// Unique object name used in the XML output.
    pub object_name: String,
    /// Ordered rows of cells that make up the table body.
    pub rows: Vec<TableRow>,
    /// Table-level formatting and layout options.
    pub options: TableOptions,
}

/// A single row of table cells.
pub type TableRow = Vec<TableCell>;

/// A single cell in a table.
#[derive(Debug, Clone)]
pub struct TableCell {
    /// Text content of the cell. Supports multiple text runs.
    pub text: Vec<CellTextRun>,
    /// Per-cell styling and layout properties.
    pub options: TableCellProps,
    /// True if this cell is a dummy placeholder for a merged area.
    pub is_merged: bool,
}

impl TableCell {
    /// Create a cell containing a single text run.
    pub fn new(text: impl Into<String>) -> Self {
        TableCell {
            text: vec![CellTextRun::new(text)],
            options: TableCellProps::default(),
            is_merged: false,
        }
    }

    /// Create an empty cell with no text content.
    pub fn empty() -> Self {
        TableCell { text: Vec::new(), options: TableCellProps::default(), is_merged: false }
    }

    /// Create a placeholder cell that represents a merged area.
    pub fn merged() -> Self {
        TableCell { text: Vec::new(), options: TableCellProps::default(), is_merged: true }
    }

    /// Set the cell background to a solid color, 6-digit hex, no `#` prefix.
    pub fn fill(mut self, c: impl Into<String>) -> Self {
        self.options.fill = Some(c.into());
        self
    }

    /// Set the cell background to a gradient fill.
    pub fn gradient_fill(mut self, g: GradientFill) -> Self {
        self.options.gradient_fill = Some(g);
        self
    }

    /// Set the font color. Hex color, no `#` prefix.
    pub fn color(mut self, c: impl Into<String>) -> Self {
        self.options.color = Some(c.into().trim_start_matches('#').to_uppercase());
        self
    }

    /// Enable bold text.
    pub fn bold(mut self) -> Self {
        self.options.bold = Some(true);
        self
    }

    /// Enable italic text.
    pub fn italic(mut self) -> Self {
        self.options.italic = Some(true);
        self
    }

    /// Enable underlined text.
    pub fn underline(mut self) -> Self {
        self.options.underline = Some(true);
        self
    }

    /// Set the font size in points.
    pub fn font_size(mut self, pt: f64) -> Self {
        self.options.font_size = Some(pt);
        self
    }

    /// Set the font face name (e.g. `"Arial"`).
    pub fn font_face(mut self, f: impl Into<String>) -> Self {
        self.options.font_face = Some(f.into());
        self
    }

    /// Set horizontal text alignment.
    pub fn align(mut self, a: crate::enums::AlignH) -> Self {
        self.options.align = Some(a);
        self
    }

    /// Set vertical text alignment.
    pub fn valign(mut self, a: crate::enums::AlignV) -> Self {
        self.options.valign = Some(a);
        self
    }

    /// Set the inner cell margin.
    pub fn margin(mut self, m: impl Into<crate::types::Margin>) -> Self {
        self.options.margin = Some(m.into());
        self
    }

    /// Set the number of columns this cell spans.
    pub fn colspan(mut self, n: u32) -> Self {
        self.options.colspan = Some(n);
        self
    }

    /// Set the number of rows this cell spans.
    pub fn rowspan(mut self, n: u32) -> Self {
        self.options.rowspan = Some(n);
        self
    }
}

/// A text run within a table cell.
#[derive(Debug, Clone)]
pub struct CellTextRun {
    /// Plain text content of this run.
    pub text: String,
    /// Bold override for this run.
    pub bold: Option<bool>,
    /// Italic override for this run.
    pub italic: Option<bool>,
    /// Underline override for this run.
    pub underline: Option<bool>,
    /// Font size in points.
    pub font_size: Option<f64>,
    /// Font family name.
    pub font_face: Option<String>,
    /// Font color as 6-digit hex, no `#` prefix.
    pub color: Option<String>,
    /// Whether a line break is inserted before this run.
    pub break_line: bool,
}

impl CellTextRun {
    /// Create a text run with the given content and default formatting.
    pub fn new(text: impl Into<String>) -> Self {
        CellTextRun {
            text: text.into(),
            bold: None,
            italic: None,
            underline: None,
            font_size: None,
            font_face: None,
            color: None,
            break_line: false,
        }
    }
}

/// Per-cell styling options.
#[derive(Debug, Clone, Default)]
pub struct TableCellProps {
    /// Number of columns this cell spans.
    pub colspan: Option<u32>,
    /// Number of rows this cell spans.
    pub rowspan: Option<u32>,
    /// Solid background color as 6-digit hex, no `#` prefix.
    pub fill: Option<String>,
    /// Gradient background fill.
    pub gradient_fill: Option<GradientFill>,
    /// Pattern fill for the cell background.
    pub pattern_fill: Option<PatternFill>,
    /// Font color as 6-digit hex, no `#` prefix.
    pub color: Option<String>,
    /// Bold text override.
    pub bold: Option<bool>,
    /// Italic text override.
    pub italic: Option<bool>,
    /// Underline text override.
    pub underline: Option<bool>,
    /// Font size in points.
    pub font_size: Option<f64>,
    /// Font family name.
    pub font_face: Option<String>,
    /// Horizontal text alignment.
    pub align: Option<AlignH>,
    /// Vertical text alignment.
    pub valign: Option<AlignV>,
    /// Inner cell margin.
    pub margin: Option<Margin>,
    /// Per-side border definitions.
    pub border: Option<CellBorders>,
    /// Optional hyperlink on the cell (URL or slide jump).
    pub hyperlink: Option<HyperlinkProps>,
}

/// Per-side borders for a table cell.
#[derive(Debug, Clone, Default)]
pub struct CellBorders {
    /// Top edge border.
    pub top: Option<BorderProps>,
    /// Right edge border.
    pub right: Option<BorderProps>,
    /// Bottom edge border.
    pub bottom: Option<BorderProps>,
    /// Left edge border.
    pub left: Option<BorderProps>,
}

/// Table-level options.
#[derive(Debug, Clone)]
pub struct TableOptions {
    /// Position and size of the table on the slide.
    pub position: PositionProps,
    /// Column widths in inches. If None, distribute evenly.
    pub col_w: Option<Vec<f64>>,
    /// Row heights in inches. If None, auto-size.
    pub row_h: Option<Vec<f64>>,
    /// Default solid background color as 6-digit hex, no `#` prefix.
    pub fill: Option<String>,
    /// Default gradient background fill.
    pub gradient_fill: Option<GradientFill>,
    /// Default font color as 6-digit hex, no `#` prefix.
    pub color: Option<String>,
    /// Default font size in points.
    pub font_size: Option<f64>,
    /// Default font family name.
    pub font_face: Option<String>,
    /// Default bold text setting.
    pub bold: Option<bool>,
    /// Default italic text setting.
    pub italic: Option<bool>,
    /// Default horizontal text alignment.
    pub align: Option<AlignH>,
    /// Default vertical text alignment.
    pub valign: Option<AlignV>,
    /// Default inner cell margin.
    pub margin: Option<Margin>,
    /// Default border applied to every cell.
    pub border: Option<BorderProps>,
    /// Drop shadow applied to the table.
    pub shadow: Option<ShadowProps>,
    /// Font size modifier for auto-page calculation.
    pub auto_page_char_weight: Option<f64>,
    /// Row height for auto-pagination (inches).
    pub auto_page_slide_starting_row_h: Option<f64>,
    /// Number of header rows to repeat on each auto-paged slide.
    pub auto_page_header_rows: Option<usize>,
    /// Object name for XML.
    pub object_name: Option<String>,
    /// Click-triggered animations on this table (each fires on its own click).
    pub animations: Vec<AnimationEffect>,
    /// OOXML built-in table style GUID, e.g. "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
    pub table_style_id: Option<String>,
}

impl Default for TableOptions {
    fn default() -> Self {
        TableOptions {
            position: PositionProps::default(),
            col_w: None,
            row_h: None,
            fill: None,
            gradient_fill: None,
            color: None,
            font_size: None,
            font_face: None,
            bold: None,
            italic: None,
            align: None,
            valign: None,
            margin: None,
            border: None,
            shadow: None,
            auto_page_char_weight: None,
            auto_page_slide_starting_row_h: None,
            auto_page_header_rows: None,
            object_name: None,
            animations: Vec::new(),
            table_style_id: None,
        }
    }
}

/// Builder for table options.
pub struct TableOptionsBuilder {
    opts: TableOptions,
}

impl TableOptionsBuilder {
    /// Create a new builder with default table options.
    pub fn new() -> Self {
        TableOptionsBuilder { opts: TableOptions::default() }
    }

    /// Set the horizontal position in inches.
    pub fn x(mut self, v: f64) -> Self { self.opts.position.x = Some(Coord::Inches(v)); self }
    /// Set the vertical position in inches.
    pub fn y(mut self, v: f64) -> Self { self.opts.position.y = Some(Coord::Inches(v)); self }
    /// Set the table width in inches.
    pub fn w(mut self, v: f64) -> Self { self.opts.position.w = Some(Coord::Inches(v)); self }
    /// Set the table height in inches.
    pub fn h(mut self, v: f64) -> Self { self.opts.position.h = Some(Coord::Inches(v)); self }
    /// Set position (x, y) in inches.
    pub fn pos(self, x: f64, y: f64) -> Self {
        self.x(x).y(y)
    }
    /// Set size (width, height) in inches.
    pub fn size(self, w: f64, h: f64) -> Self {
        self.w(w).h(h)
    }
    /// Set individual column widths in inches.
    pub fn col_w(mut self, widths: Vec<f64>) -> Self { self.opts.col_w = Some(widths); self }
    /// Set individual row heights in inches.
    pub fn row_h(mut self, heights: Vec<f64>) -> Self { self.opts.row_h = Some(heights); self }
    /// Set the default font size in points.
    pub fn font_size(mut self, pt: f64) -> Self { self.opts.font_size = Some(pt); self }
    /// Set the default font family name.
    pub fn font_face(mut self, f: impl Into<String>) -> Self { self.opts.font_face = Some(f.into()); self }
    /// Set the default background color, 6-digit hex, no `#` prefix.
    pub fn fill(mut self, c: impl Into<String>) -> Self { self.opts.fill = Some(c.into()); self }
    /// Set the default gradient background fill.
    pub fn gradient_fill(mut self, g: GradientFill) -> Self { self.opts.gradient_fill = Some(g); self }
    /// Set the default horizontal text alignment.
    pub fn align(mut self, a: AlignH) -> Self { self.opts.align = Some(a); self }
    /// Set the default vertical text alignment.
    pub fn valign(mut self, a: AlignV) -> Self { self.opts.valign = Some(a); self }
    /// Set the default cell border.
    pub fn border(mut self, b: BorderProps) -> Self { self.opts.border = Some(b); self }
    /// Set the table drop shadow.
    pub fn shadow(mut self, s: ShadowProps) -> Self { self.opts.shadow = Some(s); self }
    /// Add a click-triggered animation effect.
    pub fn animation(mut self, anim: AnimationEffect) -> Self { self.opts.animations.push(anim); self }
    /// Set the OOXML built-in table style GUID.
    pub fn table_style_id(mut self, id: impl Into<String>) -> Self { self.opts.table_style_id = Some(id.into()); self }

    /// Consume the builder and return the finished table options.
    pub fn build(self) -> TableOptions {
        self.opts
    }
}

impl Default for TableOptionsBuilder {
    fn default() -> Self {
        Self::new()
    }
}
