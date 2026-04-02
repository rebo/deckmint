use deckmint::objects::image::ImageOptionsBuilder;
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::table::{TableCell, TableOptions, TableRow};
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{Presentation, ShapeType};
use wasm_bindgen::prelude::*;

#[wasm_bindgen(start)]
pub fn start() {
    console_error_panic_hook::set_once();
}

/// WASM-exposed Presentation builder
#[wasm_bindgen]
pub struct JsPresentation {
    inner: Presentation,
}

#[wasm_bindgen]
impl JsPresentation {
    #[wasm_bindgen(constructor)]
    pub fn new() -> Self {
        JsPresentation {
            inner: Presentation::new(),
        }
    }

    /// Set presentation title
    pub fn set_title(&mut self, title: &str) {
        self.inner.title = title.to_string();
    }

    /// Set presentation author
    pub fn set_author(&mut self, author: &str) {
        self.inner.author = author.to_string();
    }

    /// Set presentation subject
    pub fn set_subject(&mut self, subject: &str) {
        self.inner.subject = subject.to_string();
    }

    /// Set presentation company
    pub fn set_company(&mut self, company: &str) {
        self.inner.company = company.to_string();
    }

    /// Add a blank slide and return its index (0-based)
    pub fn add_slide(&mut self) -> usize {
        self.inner.add_slide();
        self.inner.slide_count() - 1
    }

    /// Add a text box to a slide
    /// opts_json: JSON string with fields: x, y, w, h, fontSize, bold, italic, align, color, fill
    pub fn add_text(&mut self, slide_idx: usize, text: &str, opts_json: &str) -> Result<(), JsValue> {
        let opts = parse_text_opts(opts_json)?;
        if let Some(slide) = self.inner.slide_mut(slide_idx) {
            slide.add_text(text, opts);
        }
        Ok(())
    }

    /// Add a shape to a slide
    /// shape_type: string like "rect", "ellipse", "triangle", etc.
    /// opts_json: JSON string with fields: x, y, w, h, fill, line_color, line_width
    pub fn add_shape(&mut self, slide_idx: usize, shape_type: &str, opts_json: &str) -> Result<(), JsValue> {
        let shape = parse_shape_type(shape_type)?;
        let opts = parse_shape_opts(opts_json)?;
        if let Some(slide) = self.inner.slide_mut(slide_idx) {
            slide.add_shape(shape, opts);
        }
        Ok(())
    }

    /// Add an image from a base64-encoded string
    /// extension: "png", "jpg", "gif", "svg", etc.
    /// opts_json: JSON string with fields: x, y, w, h, alt_text, transparency
    pub fn add_image_base64(
        &mut self,
        slide_idx: usize,
        b64: &str,
        extension: &str,
        opts_json: &str,
    ) -> Result<(), JsValue> {
        let opts = parse_image_opts(opts_json)?;
        if let Some(slide) = self.inner.slide_mut(slide_idx) {
            slide.add_image_base64(b64, extension, opts)
                .map_err(|e| JsValue::from_str(&e.to_string()))?;
        }
        Ok(())
    }

    /// Add a table from a JSON array of rows
    /// rows_json: JSON array of arrays of objects: [{text, bold, italic, colspan, rowspan, fill, color, align}, ...]
    /// opts_json: JSON with: x, y, w, h, col_w (array of col widths)
    pub fn add_table(
        &mut self,
        slide_idx: usize,
        rows_json: &str,
        opts_json: &str,
    ) -> Result<(), JsValue> {
        let rows = parse_table_rows(rows_json)?;
        let opts = parse_table_opts(opts_json)?;
        if let Some(slide) = self.inner.slide_mut(slide_idx) {
            slide.add_table(rows, opts);
        }
        Ok(())
    }

    /// Set the background color of a slide (hex string, e.g. "FF0000")
    pub fn set_background_color(&mut self, slide_idx: usize, color: &str) {
        if let Some(slide) = self.inner.slide_mut(slide_idx) {
            slide.set_background_color(color);
        }
    }

    /// Add speaker notes to a slide
    pub fn add_notes(&mut self, slide_idx: usize, notes: &str) {
        if let Some(slide) = self.inner.slide_mut(slide_idx) {
            slide.add_notes(notes);
        }
    }

    /// Serialize the presentation to a Uint8Array (.pptx bytes)
    pub fn write(&self) -> Result<js_sys::Uint8Array, JsValue> {
        let bytes = self.inner.write()
            .map_err(|e| JsValue::from_str(&e.to_string()))?;
        Ok(js_sys::Uint8Array::from(bytes.as_slice()))
    }
}

// ─── JSON parsing helpers ────────────────────────────────────────────────────

fn get_f64(obj: &serde_json::Value, key: &str) -> Option<f64> {
    obj.get(key)?.as_f64()
}

fn get_str<'a>(obj: &'a serde_json::Value, key: &str) -> Option<&'a str> {
    obj.get(key)?.as_str()
}

fn get_bool(obj: &serde_json::Value, key: &str) -> Option<bool> {
    obj.get(key)?.as_bool()
}

fn parse_text_opts(json: &str) -> Result<deckmint::objects::text::TextOptions, JsValue> {
    let v: serde_json::Value = serde_json::from_str(json)
        .unwrap_or(serde_json::Value::Object(serde_json::Map::new()));

    let mut b = TextOptionsBuilder::new();
    if let Some(x) = get_f64(&v, "x") { b = b.x(x); }
    if let Some(y) = get_f64(&v, "y") { b = b.y(y); }
    if let Some(w) = get_f64(&v, "w") { b = b.w(w); }
    if let Some(h) = get_f64(&v, "h") { b = b.h(h); }
    if let Some(fs) = get_f64(&v, "fontSize") { b = b.font_size(fs); }
    if get_bool(&v, "bold") == Some(true) { b = b.bold(); }
    if get_bool(&v, "italic") == Some(true) { b = b.italic(); }
    if let Some(color) = get_str(&v, "color") { b = b.color(color); }
    if let Some(fill) = get_str(&v, "fill") { b = b.fill(fill); }
    if let Some(align) = get_str(&v, "align") {
        let a = match align {
            "center" | "ctr" => deckmint::AlignH::Center,
            "right" | "r" => deckmint::AlignH::Right,
            "justify" | "just" => deckmint::AlignH::Justify,
            _ => deckmint::AlignH::Left,
        };
        b = b.align(a);
    }
    Ok(b.build())
}

fn parse_shape_type(s: &str) -> Result<ShapeType, JsValue> {
    let st = match s.to_lowercase().as_str() {
        "rect" | "rectangle" => ShapeType::Rect,
        "ellipse" | "oval" => ShapeType::Ellipse,
        "triangle" => ShapeType::Triangle,
        "roundrect" | "round_rect" => ShapeType::RoundRect,
        "diamond" => ShapeType::Diamond,
        "pentagon" => ShapeType::Pentagon,
        "hexagon" => ShapeType::Hexagon,
        "heptagon" => ShapeType::Heptagon,
        "octagon" => ShapeType::Octagon,
        "star4" => ShapeType::Star4,
        "star5" => ShapeType::Star5,
        "star6" => ShapeType::Star6,
        "star7" => ShapeType::Star7,
        "star8" => ShapeType::Star8,
        "star10" => ShapeType::Star10,
        "star12" => ShapeType::Star12,
        "star16" => ShapeType::Star16,
        "star24" => ShapeType::Star24,
        "star32" => ShapeType::Star32,
        "arrow_right" | "rightarrow" => ShapeType::RightArrow,
        "arrow_left" | "leftarrow" => ShapeType::LeftArrow,
        "arrow_up" | "uparrow" => ShapeType::UpArrow,
        "arrow_down" | "downarrow" => ShapeType::DownArrow,
        "line" => ShapeType::Line,
        _ => return Err(JsValue::from_str(&format!("Unknown shape type: {s}"))),
    };
    Ok(st)
}

fn parse_shape_opts(json: &str) -> Result<deckmint::objects::shape::ShapeOptions, JsValue> {
    let v: serde_json::Value = serde_json::from_str(json)
        .unwrap_or(serde_json::Value::Object(serde_json::Map::new()));

    let mut b = ShapeOptionsBuilder::new();
    if let Some(x) = get_f64(&v, "x") { b = b.x(x); }
    if let Some(y) = get_f64(&v, "y") { b = b.y(y); }
    if let Some(w) = get_f64(&v, "w") { b = b.w(w); }
    if let Some(h) = get_f64(&v, "h") { b = b.h(h); }
    if let Some(fill) = get_str(&v, "fill") { b = b.fill_color(fill); }
    if get_bool(&v, "no_fill") == Some(true) { b = b.no_fill(); }
    if let Some(lc) = get_str(&v, "line_color") { b = b.line_color(lc); }
    if let Some(lw) = get_f64(&v, "line_width") { b = b.line_width(lw); }
    if let Some(r) = get_f64(&v, "rotate") { b = b.rotate(r); }
    Ok(b.build())
}

fn parse_image_opts(json: &str) -> Result<deckmint::objects::image::ImageOptions, JsValue> {
    let v: serde_json::Value = serde_json::from_str(json)
        .unwrap_or(serde_json::Value::Object(serde_json::Map::new()));

    let mut b = ImageOptionsBuilder::new();
    if let Some(x) = get_f64(&v, "x") { b = b.x(x); }
    if let Some(y) = get_f64(&v, "y") { b = b.y(y); }
    if let Some(w) = get_f64(&v, "w") { b = b.w(w); }
    if let Some(h) = get_f64(&v, "h") { b = b.h(h); }
    if let Some(alt) = get_str(&v, "alt_text") { b = b.alt_text(alt); }
    if let Some(t) = get_f64(&v, "transparency") { b = b.transparency(t); }
    if get_bool(&v, "rounding") == Some(true) { b = b.rounding(); }
    let (opts, _, _, _) = b.build();
    Ok(opts)
}

fn parse_table_rows(json: &str) -> Result<Vec<TableRow>, JsValue> {
    let arr: serde_json::Value = serde_json::from_str(json)
        .map_err(|e| JsValue::from_str(&format!("Invalid table rows JSON: {e}")))?;

    let rows_arr = arr.as_array()
        .ok_or_else(|| JsValue::from_str("rows_json must be an array"))?;

    let mut rows: Vec<TableRow> = Vec::new();
    for row_val in rows_arr {
        let cells_arr = row_val.as_array()
            .ok_or_else(|| JsValue::from_str("Each row must be an array of cells"))?;
        let mut row: TableRow = Vec::new();
        for cell_val in cells_arr {
            if cell_val.get("merge").and_then(|v| v.as_bool()) == Some(true) {
                row.push(TableCell::merged());
                continue;
            }
            let text = cell_val.get("text")
                .and_then(|v| v.as_str())
                .unwrap_or("")
                .to_string();
            let mut cell = TableCell::new(text);
            if let Some(colspan) = cell_val.get("colspan").and_then(|v| v.as_u64()) {
                cell.options.colspan = Some(colspan as u32);
            }
            if let Some(rowspan) = cell_val.get("rowspan").and_then(|v| v.as_u64()) {
                cell.options.rowspan = Some(rowspan as u32);
            }
            if let Some(fill) = cell_val.get("fill").and_then(|v| v.as_str()) {
                cell.options.fill = Some(fill.to_string());
            }
            if let Some(color) = cell_val.get("color").and_then(|v| v.as_str()) {
                cell.options.color = Some(color.to_string());
            }
            if let Some(bold) = cell_val.get("bold").and_then(|v| v.as_bool()) {
                cell.options.bold = Some(bold);
            }
            if let Some(italic) = cell_val.get("italic").and_then(|v| v.as_bool()) {
                cell.options.italic = Some(italic);
            }
            row.push(cell);
        }
        rows.push(row);
    }
    Ok(rows)
}

fn parse_table_opts(json: &str) -> Result<TableOptions, JsValue> {
    let v: serde_json::Value = serde_json::from_str(json)
        .unwrap_or(serde_json::Value::Object(serde_json::Map::new()));

    let mut b = deckmint::objects::table::TableOptionsBuilder::new();
    if let Some(x) = get_f64(&v, "x") { b = b.x(x); }
    if let Some(y) = get_f64(&v, "y") { b = b.y(y); }
    if let Some(w) = get_f64(&v, "w") { b = b.w(w); }
    if let Some(h) = get_f64(&v, "h") { b = b.h(h); }
    if let Some(col_w) = v.get("col_w").and_then(|v| v.as_array()) {
        let col_widths: Vec<f64> = col_w.iter()
            .filter_map(|v| v.as_f64())
            .collect();
        b = b.col_w(col_widths);
    }
    Ok(b.build())
}
