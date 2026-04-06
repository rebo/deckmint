use crate::enums::{EMU, ONEPT, DEF_FONT_COLOR, is_hex_color};
use crate::types::{Color, PresLayout};

/// Convert inches to EMU.
/// Values >= 100 are assumed to already be EMU and returned as-is.
pub fn inch_to_emu(inches: f64) -> i64 {
    if inches > 100.0 {
        inches as i64
    } else {
        (EMU as f64 * inches).round() as i64
    }
}

/// Convert points to EMU (using ONEPT).
pub fn val_to_pts(pt: f64) -> i64 {
    (pt * ONEPT as f64).round() as i64
}

/// Convert degrees (0-360) to PowerPoint rotation value (60000 units per degree).
pub fn convert_rotation_degrees(d: f64) -> i64 {
    let d = if d > 360.0 { d - 360.0 } else { d };
    (d * 60_000.0).round() as i64
}

/// Translate any coordinate to EMU given layout dimensions.
/// - numbers < 100: treat as inches → multiply by EMU
/// - numbers >= 100: treat as EMU already
/// - percentage strings like "75%": resolve against layout_dim
pub fn get_smart_parse_number(value: f64, is_percent: bool, pct_value: f64, layout_dim: i64) -> i64 {
    if is_percent {
        ((pct_value / 100.0) * layout_dim as f64).round() as i64
    } else if value < 100.0 {
        inch_to_emu(value)
    } else {
        value as i64
    }
}

/// Convert RGB components to hex string
pub fn rgb_to_hex(r: u8, g: u8, b: u8) -> String {
    format!("{:02X}{:02X}{:02X}", r, g, b)
}

/// Encode special XML characters
pub fn encode_xml_entities(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Generate a new UUID v4 string (uppercase, hyphenated)
pub fn new_uuid() -> String {
    uuid::Uuid::new_v4().to_string().to_uppercase()
}

/// Create a color XML element — either `<a:schemeClr>` or `<a:srgbClr>`.
/// `inner_elements` is optional XML content to nest inside the color tag.
pub fn create_color_element(color_str: &str, inner_elements: Option<&str>) -> String {
    let clean = color_str.trim_start_matches('#');
    let (is_hex, val) = if is_hex_color(clean) {
        (true, clean.to_uppercase())
    } else {
        // fall back to default if invalid
        let v = if clean.is_empty() { DEF_FONT_COLOR.to_string() } else { clean.to_string() };
        (false, v)
    };
    let tag = if is_hex { "srgbClr" } else { "schemeClr" };
    match inner_elements {
        Some(inner) if !inner.is_empty() => {
            format!("<a:{tag} val=\"{val}\">{inner}</a:{tag}>")
        }
        _ => format!("<a:{tag} val=\"{val}\"/>"),
    }
}

/// Generate `<a:solidFill>...</a:solidFill>` XML for a color string or Color.
/// Returns empty string for empty/invalid colors.
pub fn gen_xml_color_selection_str(color_str: &str, transparency: Option<f64>) -> String {
    let clean = color_str.trim_start_matches('#');
    if clean.is_empty() {
        return String::new();
    }
    let inner = match transparency {
        Some(t) if t > 0.0 => {
            let alpha = ((100.0 - t) * 1000.0).round() as i64;
            format!("<a:alpha val=\"{alpha}\"/>")
        }
        _ => String::new(),
    };
    format!("<a:solidFill>{}</a:solidFill>", create_color_element(clean, Some(&inner)))
}

/// Generate `<a:solidFill>` from a Color enum
pub fn gen_xml_color_selection(color: &Color, transparency: Option<f64>) -> String {
    match color {
        Color::Hex(h) => gen_xml_color_selection_str(h, transparency),
        Color::Theme(sc) => {
            let inner = match transparency {
                Some(t) if t > 0.0 => {
                    let alpha = ((100.0 - t) * 1000.0).round() as i64;
                    format!("<a:alpha val=\"{alpha}\"/>")
                }
                _ => String::new(),
            };
            format!("<a:solidFill>{}</a:solidFill>",
                create_color_element(sc.as_str(), Some(&inner)))
        }
        Color::ThemedWith(sc, mods) => {
            let mut inner = String::new();
            if let Some(v) = mods.lum_mod { inner.push_str(&format!("<a:lumMod val=\"{v}\"/>")); }
            if let Some(v) = mods.lum_off { inner.push_str(&format!("<a:lumOff val=\"{v}\"/>")); }
            if let Some(v) = mods.tint    { inner.push_str(&format!("<a:tint val=\"{v}\"/>")); }
            if let Some(v) = mods.shade   { inner.push_str(&format!("<a:shade val=\"{v}\"/>")); }
            if let Some(v) = mods.sat_mod { inner.push_str(&format!("<a:satMod val=\"{v}\"/>")); }
            if let Some(t) = transparency {
                if t > 0.0 {
                    let alpha = ((100.0 - t) * 1000.0).round() as i64;
                    inner.push_str(&format!("<a:alpha val=\"{alpha}\"/>"));
                }
            }
            format!("<a:solidFill>{}</a:solidFill>",
                create_color_element(sc.as_str(), Some(&inner)))
        }
    }
}

/// Generate `<a:gradFill>` XML from a [`crate::types::GradientFill`] definition.
pub fn gen_xml_grad_fill(grad: &crate::types::GradientFill) -> String {
    use crate::types::GradientType;
    let mut s = String::from("<a:gradFill rotWithShape=\"1\"><a:gsLst>");
    for stop in &grad.stops {
        let pos = (stop.position * 1000.0).round() as i64;
        // Use the full Color-aware renderer which handles Hex, Theme, and ThemedWith
        let color_xml = gen_xml_color_selection(&stop.color, stop.transparency);
        // gen_xml_color_selection wraps in <a:solidFill>; strip that wrapper for <a:gs>
        let inner_xml = color_xml
            .strip_prefix("<a:solidFill>").unwrap_or(&color_xml)
            .strip_suffix("</a:solidFill>").unwrap_or(&color_xml);
        s.push_str(&format!("<a:gs pos=\"{pos}\">{inner_xml}</a:gs>"));
    }
    s.push_str("</a:gsLst>");
    match grad.gradient_type {
        GradientType::Linear => {
            let ang = (grad.angle * 60_000.0).round() as i64;
            s.push_str(&format!("<a:lin ang=\"{ang}\" scaled=\"0\"/>"));
        }
        GradientType::Radial => {
            s.push_str("<a:path path=\"circle\"><a:fillToRect l=\"50000\" t=\"50000\" r=\"50000\" b=\"50000\"/></a:path>");
        }
    }
    s.push_str("</a:gradFill>");
    s
}

/// Create `<a:glow>` element for text glow effect
pub fn create_glow_element(size: f64, color: &str, opacity: f64) -> String {
    let rad = val_to_pts(size);
    // Normalise: if caller passed a percentage (e.g. 75.0 meaning 75%), convert to 0‑1.
    let norm = if opacity > 1.0 { opacity / 100.0 } else { opacity };
    let alpha = (norm.clamp(0.0, 1.0) * 100_000.0).round() as i64;
    let inner = format!("<a:alpha val=\"{alpha}\"/>");
    format!("<a:glow rad=\"{rad}\">{}</a:glow>", create_color_element(color, Some(&inner)))
}

/// Get the next relationship ID for a slide
/// (equivalent to `getNewRelId` in TS: rels.len + relsChart.len + relsMedia.len + 1)
pub fn get_new_rel_id(rels_count: usize, rels_chart_count: usize, rels_media_count: usize) -> u32 {
    (rels_count + rels_chart_count + rels_media_count + 1) as u32
}

/// Resolve a coordinate value to EMU.
/// `is_x` determines whether to use layout width or height for percentage resolution.
pub fn coord_to_emu(value: f64, is_percent: bool, pct_value: f64, is_x: bool, layout: &PresLayout) -> i64 {
    let layout_dim = if is_x { layout.width } else { layout.height };
    get_smart_parse_number(value, is_percent, pct_value, layout_dim)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_inch_to_emu() {
        assert_eq!(inch_to_emu(1.0), 914_400);
        assert_eq!(inch_to_emu(0.5), 457_200);
        // values >= 100 pass through
        assert_eq!(inch_to_emu(914_400.0), 914_400);
    }

    #[test]
    fn test_val_to_pts() {
        assert_eq!(val_to_pts(1.0), 12_700);
        assert_eq!(val_to_pts(12.0), 152_400);
    }

    #[test]
    fn test_convert_rotation() {
        assert_eq!(convert_rotation_degrees(90.0), 5_400_000);
        assert_eq!(convert_rotation_degrees(270.0), 16_200_000);
    }

    #[test]
    fn test_encode_xml_entities() {
        assert_eq!(encode_xml_entities("a & b < c > d"), "a &amp; b &lt; c &gt; d");
        assert_eq!(encode_xml_entities("say \"hi\""), "say &quot;hi&quot;");
    }

    #[test]
    fn test_create_color_element_hex() {
        assert_eq!(create_color_element("FF0000", None), "<a:srgbClr val=\"FF0000\"/>");
        assert_eq!(
            create_color_element("FF0000", Some("<a:alpha val=\"75000\"/>")),
            "<a:srgbClr val=\"FF0000\"><a:alpha val=\"75000\"/></a:srgbClr>"
        );
    }

    #[test]
    fn test_create_color_element_scheme() {
        assert_eq!(create_color_element("accent1", None), "<a:schemeClr val=\"accent1\"/>");
    }

    #[test]
    fn test_gen_xml_color_selection() {
        let result = gen_xml_color_selection_str("FF0000", None);
        assert_eq!(result, "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>");
    }
}
