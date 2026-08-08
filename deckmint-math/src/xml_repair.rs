// xml_repair.rs — string-level repairs for `latex2mathml` 0.2.3's MathML
// serialisation bugs.
//
// `latex2mathml` tokenises and parses LaTeX correctly for these cases, but its
// `Display` impls for the resulting AST nodes emit malformed MathML text:
//
// 1. `Node::Operator` / `Node::Letter` (and friends) write the literal
//    character straight into MathML text content with no XML escaping, so a
//    bare `<` (e.g. `x < 3`) or `&` (e.g. `A \& B`) in the source LaTeX
//    produces `<mo><</mo>` — invalid XML that a conformant parser (this
//    crate's own `mathml_to_omml::convert`, built on `quick-xml`) rejects.
// 2. `Node::SizedParen`'s `Display` impl has a typo: it opens
//    `<mo maxsize=… minsize=…>` but closes with `</mro>` instead of `</mo>`.
//    `</mro>` never appears in valid MathML output any other way, so a blind
//    string replace is safe.
//
// Both are one-line bugs in `latex2mathml` itself (an external crates.io
// dependency, currently pinned to `"0.2"`); until they're fixed upstream,
// `latex_to_omml` repairs the MathML string here before handing it to
// `mathml_to_omml::convert`. Ported faithfully from the same fix implemented
// in the revision_atlas `deckgen` crate's `src/maths.rs`, which discovered
// both bugs during a 10,348-span corpus sweep (see that repo's
// `wayfinder/deckgen/assets/maths-sweep.md`).

/// Escape any `<` not immediately followed by a tag-start character
/// (`[a-zA-Z/?!]`) to `&lt;`, and any `&` not already starting a well-formed
/// entity/char reference (`&name;`, `&#123;`, `&#x1F;`) to `&amp;`.
///
/// No-op on already-well-formed MathML, so it is safe to apply
/// unconditionally to every `latex2mathml` output before parsing it.
pub(crate) fn escape_bad_xml_chars(mathml: &str) -> String {
    let chars: Vec<char> = mathml.chars().collect();
    let mut out = String::with_capacity(mathml.len() + 16);
    let mut i = 0usize;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '<' => {
                let next = chars.get(i + 1).copied();
                let is_tag_start =
                    matches!(next, Some(n) if n.is_ascii_alphabetic() || n == '/' || n == '?' || n == '!');
                out.push_str(if is_tag_start { "<" } else { "&lt;" });
            }
            '&' => {
                out.push_str(if is_wellformed_entity(&chars, i) { "&" } else { "&amp;" });
            }
            _ => out.push(c),
        }
        i += 1;
    }
    out
}

/// Does `chars[amp_idx]` (an `&`) begin a well-formed `&name;` / `&#123;` /
/// `&#x1F;` reference?
fn is_wellformed_entity(chars: &[char], amp_idx: usize) -> bool {
    let mut j = amp_idx + 1;
    if chars.get(j) == Some(&'#') {
        j += 1;
        if matches!(chars.get(j), Some('x') | Some('X')) {
            j += 1;
            let start = j;
            while chars.get(j).map(|c| c.is_ascii_hexdigit()).unwrap_or(false) {
                j += 1;
            }
            j > start && chars.get(j) == Some(&';')
        } else {
            let start = j;
            while chars.get(j).map(|c| c.is_ascii_digit()).unwrap_or(false) {
                j += 1;
            }
            j > start && chars.get(j) == Some(&';')
        }
    } else {
        let start = j;
        while chars.get(j).map(|c| c.is_ascii_alphanumeric()).unwrap_or(false) {
            j += 1;
        }
        j > start && chars.get(j) == Some(&';')
    }
}

/// `latex2mathml` 0.2.3's `Node::SizedParen` `Display` impl has a typo — opens
/// `<mo maxsize=… minsize=…>` but closes with `</mro>` instead of `</mo>`.
/// `</mro>` never appears in valid output any other way, so a blind replace is
/// safe.
pub(crate) fn fix_sizedparen_typo(mathml: &str) -> String {
    mathml.replace("</mro>", "</mo>")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn escape_bad_xml_chars_is_noop_on_wellformed_mathml() {
        let input = r#"<math><mrow><mi>x</mi><mo>&lt;</mo><mn>3</mn></mrow></math>"#;
        assert_eq!(escape_bad_xml_chars(input), input);
    }

    #[test]
    fn escape_bad_xml_chars_repairs_raw_angle_bracket() {
        let input = "<mo><</mo>";
        assert_eq!(escape_bad_xml_chars(input), "<mo>&lt;</mo>");
    }

    #[test]
    fn escape_bad_xml_chars_repairs_raw_ampersand() {
        let input = "<mi>A</mi><mo>&</mo><mi>B</mi>";
        assert_eq!(escape_bad_xml_chars(input), "<mi>A</mi><mo>&amp;</mo><mi>B</mi>");
    }

    #[test]
    fn fix_sizedparen_typo_repairs_mro_close_tag() {
        let input = r#"<mrow><mo maxsize="2">(</mro></mrow>"#;
        assert_eq!(fix_sizedparen_typo(input), r#"<mrow><mo maxsize="2">(</mo></mrow>"#);
    }
}
