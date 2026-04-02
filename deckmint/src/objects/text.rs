use crate::enums::{AlignH, AlignV};
use crate::types::{AnimationEffect, Coord, FieldType, GlowProps, GradientFill, HyperlinkProps, Margin, PositionProps, ShadowProps, ShapeLineProps, TextOutlineProps};

/// A fully resolved text object placed on a slide.
#[derive(Debug, Clone)]
pub struct TextObject {
    /// Internal object name used for relationship and shape identification.
    pub object_name: String,
    /// Ordered list of text runs that make up the content of this text box.
    pub text: Vec<TextRun>,
    /// Layout and paragraph-level formatting options for this text box.
    pub options: TextOptions,
}

/// A single run of text with optional per-run formatting.
#[derive(Debug, Clone)]
pub struct TextRun {
    /// The text content of this run.
    pub text: String,
    /// Character-level formatting options for this run.
    pub options: TextRunOptions,
    /// True if a line-break should follow this run.
    pub break_line: bool,
    /// True if a soft line-break (`<a:br>`) should precede this run.
    pub soft_break_before: bool,
    /// When set, this run renders as an `<a:fld>` field (e.g. slide number) instead of `<a:r>`.
    pub field: Option<FieldType>,
    /// Pre-rendered OMML equation XML (the `<a14:m>` element). When set, this run
    /// renders as an inline equation instead of a normal text run.
    pub equation_omml: Option<String>,
}

impl TextRun {
    /// Create a new text run with the given content and default formatting.
    pub fn new(text: impl Into<String>) -> Self {
        TextRun { text: text.into(), options: TextRunOptions::default(), break_line: false, soft_break_before: false, field: None, equation_omml: None }
    }

    /// Create a text run containing a LaTeX equation.
    ///
    /// The equation is converted to native OMML (editable in PowerPoint).
    /// Requires the `math` feature.
    ///
    /// ```rust,no_run
    /// use deckmint::objects::text::TextRun;
    /// let run = TextRun::equation(r"x^2 + 3x - 7").unwrap();
    /// ```
    #[cfg(feature = "math")]
    pub fn equation(latex: &str) -> Result<Self, crate::error::PptxError> {
        let omml = deckmint_math::latex_to_omml(latex)
            .map_err(|e| crate::error::PptxError::InvalidArgument(
                format!("equation conversion failed: {e}"),
            ))?;
        Ok(TextRun {
            text: String::new(),
            options: TextRunOptions::default(),
            break_line: false,
            soft_break_before: false,
            field: None,
            equation_omml: Some(omml),
        })
    }
}

/// Per-run formatting (character-level).
#[derive(Debug, Clone, Default)]
pub struct TextRunOptions {
    /// Whether the text is bold.
    pub bold: Option<bool>,
    /// Whether the text is italic.
    pub italic: Option<bool>,
    /// OOXML underline style: "sng", "dbl", "dash", "dashHeavy", "dashLong", "dashLongHeavy",
    /// "dotDash", "dotDashHeavy", "dotDotDash", "dotDotDashHeavy", "dotted", "heavy",
    /// "wavy", "wavyDbl", "wavyHeavy". Use "sng" for basic underline.
    pub underline: Option<String>,
    /// Underline color as 6-digit hex, no `#` prefix.
    pub underline_color: Option<String>,
    /// Strike style: "sng" (single) or "dbl" (double). Use `TextRunBuilder::strike()` for single.
    pub strike: Option<String>,
    /// Font size in points.
    pub font_size: Option<f64>,
    /// Font face name, e.g. "Arial" or "Calibri".
    pub font_face: Option<String>,
    /// Font color as 6-digit hex, no `#` prefix.
    pub color: Option<String>,
    /// Text transparency 0–100 (0 = opaque, 100 = fully transparent).
    pub transparency: Option<f64>,
    /// Highlight color as 6-digit hex (no #). Rendered as `<a:highlight>`.
    pub highlight: Option<String>,
    /// Character spacing in points (positive expands, negative condenses).
    pub char_spacing: Option<f64>,
    /// Whether this run is superscript.
    pub superscript: bool,
    /// Whether this run is subscript.
    pub subscript: bool,
    /// Language tag for spell-checking, e.g. "en-US".
    pub lang: Option<String>,
    /// Hyperlink attached to this text run.
    pub hyperlink: Option<HyperlinkProps>,
    /// Glow effect around the text run.
    pub glow: Option<GlowProps>,
    /// Outline (stroke) around the text run.
    pub outline: Option<TextOutlineProps>,
}

/// A tab stop in a paragraph.
#[derive(Debug, Clone)]
pub struct TabStop {
    /// Position in inches from left margin.
    pub pos_inches: f64,
    /// Alignment: "l" (left), "ctr" (center), "r" (right), "dec" (decimal).
    pub align: String,
}

impl TabStop {
    /// Create a tab stop at the given position with the given alignment.
    ///
    /// `align` is one of: `"l"` (left), `"ctr"` (center), `"r"` (right), `"dec"` (decimal).
    pub fn new(pos_inches: f64, align: impl Into<String>) -> Self {
        TabStop { pos_inches, align: align.into() }
    }
}

/// Paragraph and layout options for a text box.
#[derive(Debug, Clone)]
pub struct TextOptions {
    /// Position and size of the text box on the slide.
    pub position: PositionProps,
    /// Horizontal text alignment within the text box.
    pub align: Option<AlignH>,
    /// Vertical text alignment within the text box.
    pub valign: Option<AlignV>,
    /// Internal margin (inset) of the text box in inches.
    pub margin: Option<Margin>,
    /// Default font size in points for all runs that do not specify their own.
    pub font_size: Option<f64>,
    /// Default font face name for all runs that do not specify their own.
    pub font_face: Option<String>,
    /// Default font color as 6-digit hex, no `#` prefix.
    pub color: Option<String>,
    /// Default bold setting for all runs that do not specify their own.
    pub bold: Option<bool>,
    /// Default italic setting for all runs that do not specify their own.
    pub italic: Option<bool>,
    /// Whether text flows right-to-left.
    pub rtl_mode: bool,
    /// Line spacing in points (fixed spacing).
    pub line_spacing: Option<f64>,
    /// Line spacing as a multiple of normal line height, e.g. 1.5.
    pub line_spacing_multiple: Option<f64>,
    /// Space before each paragraph in points.
    pub para_space_before: Option<f64>,
    /// Space after each paragraph in points.
    pub para_space_after: Option<f64>,
    /// Bullet or numbering properties for paragraphs.
    pub bullet: Option<BulletProps>,
    /// Drop shadow effect on the text box.
    pub shadow: Option<ShadowProps>,
    /// Solid background fill color as 6-digit hex, no `#` prefix.
    pub fill: Option<String>,
    /// Gradient background fill for the text box.
    pub gradient_fill: Option<GradientFill>,
    /// Text auto-fit behavior for the text box.
    pub fit: Option<TextFit>,
    /// Whether text wraps within the text box.
    pub wrap: Option<bool>,
    /// Text direction: "horz", "vert", "vert270", "wordArtVert", "mongolianVert", "eaVert"
    pub vert: Option<String>,
    /// Paragraph indent level (0-based).
    pub indent_level: Option<u32>,
    /// Tab stops for paragraph layout.
    pub tab_stops: Option<Vec<TabStop>>,
    /// Internal body properties (computed from margin, valign, etc.).
    pub body_prop: Option<BodyProp>,
    /// Rotation in degrees (clockwise).
    pub rotate: Option<f64>,
    /// Flip horizontally.
    pub flip_h: bool,
    /// Flip vertically.
    pub flip_v: bool,
    /// Border line around the text box.
    pub line: Option<ShapeLineProps>,
    /// Number of text columns (≥ 2 to enable multi-column layout).
    pub num_columns: Option<u32>,
    /// Spacing between columns in inches (only used when num_columns ≥ 2).
    pub column_spacing: Option<f64>,
    /// Click-triggered animations on this text box (each fires on its own click).
    /// Use `.animation()` builder method to append; supports multiple entries with
    /// different `TextTarget` values to animate different character/paragraph ranges.
    pub animations: Vec<AnimationEffect>,
}

/// Text auto-fit behavior for a text box.
#[derive(Debug, Clone, PartialEq)]
pub enum TextFit {
    /// Do not auto-fit text.
    None,
    /// Shrink text font size to fit within the fixed text box.
    Shrink,
    /// Resize the text box to fit its content.
    Resize,
}

/// Internal body properties for a text box, mapped to `<a:bodyPr>`.
#[derive(Debug, Clone, Default)]
pub struct BodyProp {
    /// Whether text wraps inside the text box.
    pub wrap: bool,
    /// Left inset in EMUs.
    pub l_ins: Option<i64>,
    /// Top inset in EMUs.
    pub t_ins: Option<i64>,
    /// Right inset in EMUs.
    pub r_ins: Option<i64>,
    /// Bottom inset in EMUs.
    pub b_ins: Option<i64>,
    /// Vertical anchor: "t" (top), "ctr" (center), "b" (bottom).
    pub anchor: Option<String>,
    /// Text direction for the body, e.g. "vert", "vert270".
    pub vert: Option<String>,
    /// Whether auto-fit is enabled.
    pub auto_fit: bool,
}

/// Bullet or numbering properties for a paragraph.
#[derive(Debug, Clone)]
pub struct BulletProps {
    /// The type of bullet to render.
    pub bullet_type: BulletType,
    /// Unicode character code for custom bullet characters.
    pub character_code: Option<String>,
    /// Hanging indent in inches for bullet text.
    pub indent: Option<f64>,
    /// Starting number for numbered bullets.
    pub number_start_at: Option<u32>,
    /// Numbering style, e.g. "arabicPeriod", "romanUcPeriod".
    pub style: Option<String>,
}

/// The type of bullet used in a paragraph.
#[derive(Debug, Clone, PartialEq)]
pub enum BulletType {
    /// Standard filled-circle bullet.
    Default,
    /// Auto-incrementing numbered bullet.
    Numbered,
    /// Custom Unicode character bullet.
    Character,
}

impl Default for TextOptions {
    fn default() -> Self {
        TextOptions {
            position: PositionProps::default(),
            align: None,
            valign: None,
            margin: None,
            font_size: None,
            font_face: None,
            color: None,
            bold: None,
            italic: None,
            rtl_mode: false,
            line_spacing: None,
            line_spacing_multiple: None,
            para_space_before: None,
            para_space_after: None,
            bullet: None,
            shadow: None,
            fill: None,
            gradient_fill: None,
            fit: None,
            wrap: Some(true),
            vert: None,
            indent_level: None,
            tab_stops: None,
            body_prop: None,
            rotate: None,
            flip_h: false,
            flip_v: false,
            line: None,
            num_columns: None,
            column_spacing: None,
            animations: Vec::new(),
        }
    }
}

/// Builder for a single `TextRun` (rich-text character run).
///
/// ```rust,no_run
/// use deckmint::objects::text::TextRunBuilder;
/// let run = TextRunBuilder::new("Hello").color("FF0000").bold().font_size(24.0).build();
/// ```
pub struct TextRunBuilder {
    run: TextRun,
}

impl TextRunBuilder {
    /// Create a new text run builder with the given text content.
    pub fn new(text: impl Into<String>) -> Self {
        TextRunBuilder { run: TextRun::new(text) }
    }

    /// Create a text run builder for a LaTeX equation.
    ///
    /// The equation is converted to native OMML (editable in PowerPoint).
    /// Requires the `math` feature.
    ///
    /// ```rust,no_run
    /// use deckmint::TextRunBuilder;
    /// let run = TextRunBuilder::equation(r"\frac{a}{b}").unwrap().build();
    /// ```
    #[cfg(feature = "math")]
    pub fn equation(latex: &str) -> Result<Self, crate::error::PptxError> {
        let run = TextRun::equation(latex)?;
        Ok(TextRunBuilder { run })
    }

    /// Set the font color as 6-digit hex, no `#` prefix.
    pub fn color(mut self, c: impl Into<String>) -> Self {
        self.run.options.color = Some(c.into().trim_start_matches('#').to_uppercase());
        self
    }
    /// Enable bold formatting.
    pub fn bold(mut self) -> Self { self.run.options.bold = Some(true); self }
    /// Enable italic formatting.
    pub fn italic(mut self) -> Self { self.run.options.italic = Some(true); self }
    /// Set the font size in points.
    pub fn font_size(mut self, pt: f64) -> Self { self.run.options.font_size = Some(pt); self }
    /// Set the font face name, e.g. "Arial".
    pub fn font_face(mut self, f: impl Into<String>) -> Self { self.run.options.font_face = Some(f.into()); self }
    /// Set the underline style, e.g. "sng" for single underline.
    pub fn underline(mut self, style: impl Into<String>) -> Self { self.run.options.underline = Some(style.into()); self }
    /// Set the underline color as 6-digit hex, no `#` prefix.
    pub fn underline_color(mut self, c: impl Into<String>) -> Self { self.run.options.underline_color = Some(c.into().trim_start_matches('#').to_uppercase()); self }
    /// Apply single strikethrough.
    pub fn strike(mut self) -> Self { self.run.options.strike = Some("sng".to_string()); self }
    /// Apply double strikethrough.
    pub fn strike_double(mut self) -> Self { self.run.options.strike = Some("dbl".to_string()); self }
    /// Set text transparency, 0.0–100.0 (0 = opaque, 100 = fully transparent).
    pub fn transparency(mut self, t: f64) -> Self { self.run.options.transparency = Some(t); self }
    /// Format this run as superscript.
    pub fn superscript(mut self) -> Self { self.run.options.superscript = true; self }
    /// Format this run as subscript.
    pub fn subscript(mut self) -> Self { self.run.options.subscript = true; self }
    /// Set the highlight (background) color as 6-digit hex, no `#` prefix.
    pub fn highlight(mut self, c: impl Into<String>) -> Self { self.run.options.highlight = Some(c.into().trim_start_matches('#').to_uppercase()); self }
    /// Set character spacing in points.
    pub fn char_spacing(mut self, v: f64) -> Self { self.run.options.char_spacing = Some(v); self }
    /// Set the language tag for spell-checking, e.g. "en-US".
    pub fn lang(mut self, l: impl Into<String>) -> Self { self.run.options.lang = Some(l.into()); self }
    /// End this run with a paragraph break (starts a new `<a:p>`).
    pub fn break_line(mut self) -> Self { self.run.break_line = true; self }
    /// Insert a soft line break (`<a:br>`) before this run (same paragraph).
    pub fn soft_break_before(mut self) -> Self { self.run.soft_break_before = true; self }
    /// Attach a hyperlink to this text run.
    pub fn hyperlink(mut self, hl: HyperlinkProps) -> Self { self.run.options.hyperlink = Some(hl); self }
    /// Apply a glow effect around this text run.
    pub fn glow(mut self, g: GlowProps) -> Self { self.run.options.glow = Some(g); self }
    /// Apply an outline (stroke) around this text run.
    pub fn outline(mut self, o: TextOutlineProps) -> Self { self.run.options.outline = Some(o); self }
    /// Make this run an auto-updating field (slide number, date, etc.) instead of static text.
    pub fn field(mut self, ft: FieldType) -> Self { self.run.field = Some(ft); self }

    /// Consume the builder and return the finished text run.
    pub fn build(self) -> TextRun {
        self.run
    }
}

/// Builder for paragraph and layout options of a text box.
pub struct TextOptionsBuilder {
    opts: TextOptions,
}

impl TextOptionsBuilder {
    /// Create a new text options builder with default values.
    pub fn new() -> Self {
        TextOptionsBuilder { opts: TextOptions::default() }
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
    /// Set the default font size in points.
    pub fn font_size(mut self, pt: f64) -> Self { self.opts.font_size = Some(pt); self }
    /// Set the default font face name, e.g. "Arial".
    pub fn font_face(mut self, f: impl Into<String>) -> Self { self.opts.font_face = Some(f.into()); self }
    /// Set the default font color as 6-digit hex, no `#` prefix.
    pub fn color(mut self, c: impl Into<String>) -> Self { self.opts.color = Some(c.into()); self }
    /// Enable bold formatting as the default for all runs.
    pub fn bold(mut self) -> Self { self.opts.bold = Some(true); self }
    /// Enable italic formatting as the default for all runs.
    pub fn italic(mut self) -> Self { self.opts.italic = Some(true); self }
    /// Set horizontal text alignment.
    pub fn align(mut self, a: AlignH) -> Self { self.opts.align = Some(a); self }
    /// Set vertical text alignment.
    pub fn valign(mut self, a: AlignV) -> Self { self.opts.valign = Some(a); self }
    /// Enable right-to-left text direction.
    pub fn rtl(mut self) -> Self { self.opts.rtl_mode = true; self }
    /// Set the solid background fill color as 6-digit hex, no `#` prefix.
    pub fn fill(mut self, c: impl Into<String>) -> Self { self.opts.fill = Some(c.into()); self }
    /// Set a gradient background fill for the text box.
    pub fn gradient_fill(mut self, g: GradientFill) -> Self { self.opts.gradient_fill = Some(g); self }
    /// Set the internal margin (inset) of the text box.
    pub fn margin(mut self, m: impl Into<Margin>) -> Self { self.opts.margin = Some(m.into()); self }
    /// Apply a drop shadow effect to the text box.
    pub fn shadow(mut self, s: ShadowProps) -> Self { self.opts.shadow = Some(s); self }
    /// Set fixed line spacing in points.
    pub fn line_spacing(mut self, v: f64) -> Self { self.opts.line_spacing = Some(v); self }
    /// Set line spacing as a multiple (e.g. `1.5` for 150% line spacing).
    pub fn line_spacing_multiple(mut self, mult: f64) -> Self { self.opts.line_spacing_multiple = Some(mult); self }
    /// Set the paragraph space before in points.
    pub fn para_space_before(mut self, pt: f64) -> Self { self.opts.para_space_before = Some(pt); self }
    /// Set the paragraph space after in points.
    pub fn para_space_after(mut self, pt: f64) -> Self { self.opts.para_space_after = Some(pt); self }
    /// Set the paragraph indent level (0-based).
    pub fn indent_level(mut self, level: u32) -> Self { self.opts.indent_level = Some(level); self }
    /// Set the text auto-fit behavior.
    pub fn fit(mut self, f: TextFit) -> Self { self.opts.fit = Some(f); self }
    /// Resize the text box to fit its content (`<a:spAutoFit/>`).
    pub fn autofit(mut self) -> Self { self.opts.fit = Some(TextFit::Resize); self }
    /// Shrink text to fit the fixed text box (`<a:normAutofit/>`).
    pub fn shrink_text(mut self) -> Self { self.opts.fit = Some(TextFit::Shrink); self }
    /// Set text direction. Values: "horz" (default), "vert", "vert270", "wordArtVert",
    /// "mongolianVert", "eaVert"
    pub fn text_direction(mut self, v: impl Into<String>) -> Self { self.opts.vert = Some(v.into()); self }
    /// Set tab stops for this text box.
    pub fn tab_stops(mut self, stops: Vec<TabStop>) -> Self { self.opts.tab_stops = Some(stops); self }
    /// Set rotation in degrees (clockwise).
    pub fn rotate(mut self, deg: f64) -> Self { self.opts.rotate = Some(deg); self }
    /// Flip the text box horizontally.
    pub fn flip_h(mut self) -> Self { self.opts.flip_h = true; self }
    /// Flip the text box vertically.
    pub fn flip_v(mut self) -> Self { self.opts.flip_v = true; self }
    /// Add a border line around the text box.
    pub fn line(mut self, l: ShapeLineProps) -> Self { self.opts.line = Some(l); self }
    /// Set the border color as 6-digit hex, no `#` prefix, creating a line if needed.
    pub fn line_color(mut self, color: impl Into<String>) -> Self {
        let line = self.opts.line.get_or_insert_with(ShapeLineProps::default);
        line.color = Some(color.into().trim_start_matches('#').to_uppercase());
        self
    }
    /// Set the border width in points, creating a line if needed.
    pub fn line_width(mut self, pt: f64) -> Self {
        let line = self.opts.line.get_or_insert_with(ShapeLineProps::default);
        line.width = Some(pt);
        self
    }
    /// Flow text across `n` columns.
    pub fn columns(mut self, n: u32) -> Self { self.opts.num_columns = Some(n); self }
    /// Gap between columns in inches (only meaningful when columns ≥ 2).
    pub fn column_spacing(mut self, inches: f64) -> Self { self.opts.column_spacing = Some(inches); self }
    /// Append a click-triggered animation effect to this text box.
    pub fn animation(mut self, anim: AnimationEffect) -> Self { self.opts.animations.push(anim); self }

    /// Attach a colour-emphasis animation that targets the character range covered by
    /// run `run_idx` (0-based index into the `runs` slice you will pass to `add_text_runs`).
    /// Character offsets are computed automatically from the run text lengths.
    ///
    /// ```rust,no_run
    /// # use deckmint::objects::text::{TextOptionsBuilder, TextRunBuilder};
    /// # use deckmint::types::AnimationEffect;
    /// let runs = vec![
    ///     TextRunBuilder::new("First sentence.").font_size(24.0).build(),
    ///     TextRunBuilder::new(" Second sentence.").font_size(24.0).build(),
    /// ];
    /// let opts = TextOptionsBuilder::new().x(1.0).y(2.0).w(8.0).h(1.0)
    ///     .animation_on_run(AnimationEffect::font_color("FF0000"), &runs, 0)
    ///     .animation_on_run(AnimationEffect::font_color("0000FF"), &runs, 1)
    ///     .build();
    /// ```
    pub fn animation_on_run(mut self, anim: AnimationEffect, runs: &[TextRun], run_idx: usize) -> Self {
        let (st, end) = char_range_for_run(runs, run_idx);
        self.opts.animations.push(anim.with_char_range(st, end));
        self
    }

    /// Consume the builder and return the finished text options.
    pub fn build(self) -> TextOptions {
        self.opts
    }
}

impl Default for TextOptionsBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Returns the `(st, end)` character-index range (both inclusive, 0-based) that corresponds
/// to the text of `runs[run_idx]` within the full concatenated paragraph string.
///
/// Use this with `AnimationEffect::font_color(…).with_char_range(st, end)` to target a
/// specific run inside a `add_text_runs` paragraph, or use the convenience method
/// `TextOptionsBuilder::animation_on_run(anim, runs, run_idx)` which calls this internally.
///
/// Panics if `run_idx >= runs.len()` or if the run has zero characters.
pub fn char_range_for_run(runs: &[TextRun], run_idx: usize) -> (u32, u32) {
    let st: usize = runs[..run_idx].iter().map(|r| r.text.chars().count()).sum();
    let run_len = runs[run_idx].text.chars().count();
    assert!(run_len > 0, "char_range_for_run: run {run_idx} has zero characters");
    let end = st + run_len - 1;
    (st as u32, end as u32)
}
