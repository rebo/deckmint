//! Comprehensive text formatting: underline styles, glow, outline, highlight,
//! superscript, subscript, strikethrough, character spacing, text fit modes,
//! vertical text direction, and RTL.

use deckmint::objects::text::{TextFit, TextOptionsBuilder, TextRunBuilder};
use deckmint::types::TextOutlineProps;
use deckmint::{AlignH, AlignV, GlowProps, Presentation};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Text Effects Showcase".to_string();

    // ══════════════════════════════════════════════════════════
    // Slide 1: All 15 underline styles
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#FAFAFA");

        s.add_text(
            "Underline Styles",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(28.0)
                .bold()
                .color("#1B2A4A")
                .build(),
        );

        // Divider line
        s.add_shape(
            deckmint::ShapeType::Rect,
            deckmint::objects::shape::ShapeOptionsBuilder::new()
                .bounds(0.5, 0.8, 9.0, 0.03)
                .fill_color("#4472C4")
                .build(),
        );

        let underline_styles = [
            ("Single", "sng"),
            ("Double", "dbl"),
            ("Dash", "dash"),
            ("Dotted", "dotted"),
            ("Wavy", "wavy"),
            ("Heavy", "heavy"),
            ("Dash Heavy", "dashHeavy"),
            ("Dotted Heavy", "dottedHeavy"),
            ("Wavy Heavy", "wavyHeavy"),
            ("Wavy Double", "wavyDbl"),
            ("Dash Long", "dashLong"),
            ("Dash Long Heavy", "dashLongHeavy"),
            ("Dot Dash", "dotDash"),
            ("Dot Dash Heavy", "dotDashHeavy"),
        ];

        let cols = 3;
        let col_w = 2.8;
        let row_h = 0.55;
        let x_start = 0.6;
        let y_start = 1.1;

        for (i, (label, style)) in underline_styles.iter().enumerate() {
            let col = i % cols;
            let row = i / cols;
            let x = x_start + col as f64 * (col_w + 0.15);
            let y = y_start + row as f64 * row_h;

            let mut run = TextRunBuilder::new(*label)
                .font_size(15.0)
                .color("#333333")
                .build();
            run.options.underline = Some(style.to_string());

            s.add_text_runs(
                vec![run],
                TextOptionsBuilder::new()
                    .bounds(x, y, col_w, 0.45)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Colored underline demo at the bottom
        s.add_text_runs(
            vec![
                TextRunBuilder::new("Colored underline: ")
                    .font_size(16.0)
                    .color("#555555")
                    .build(),
                TextRunBuilder::new("Red underline")
                    .font_size(16.0)
                    .underline("sng")
                    .underline_color("#FF0000")
                    .color("#333333")
                    .build(),
                TextRunBuilder::new("  |  ")
                    .font_size(16.0)
                    .color("#999999")
                    .build(),
                TextRunBuilder::new("Blue wavy")
                    .font_size(16.0)
                    .underline("wavy")
                    .underline_color("#4472C4")
                    .color("#333333")
                    .build(),
                TextRunBuilder::new("  |  ")
                    .font_size(16.0)
                    .color("#999999")
                    .build(),
                TextRunBuilder::new("Green dashed")
                    .font_size(16.0)
                    .underline("dash")
                    .underline_color("#70AD47")
                    .color("#333333")
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(0.5, 4.6, 9.0, 0.7)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Text glow, outline, highlight effects
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#1B2A4A");

        s.add_text(
            "Text Effects: Glow, Outline & Highlight",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(26.0)
                .bold()
                .color("#FFFFFF")
                .build(),
        );

        // Glow effects row
        s.add_text(
            "GLOW EFFECTS",
            TextOptionsBuilder::new()
                .bounds(0.5, 1.0, 9.0, 0.4)
                .font_size(12.0)
                .color("#8FAADC")
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Gold Glow")
                    .font_size(32.0)
                    .color("#FFFFFF")
                    .bold()
                    .glow(GlowProps::new(10.0, "#FFC000", 0.8))
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(0.5, 1.4, 4.0, 0.8)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Cyan Glow")
                    .font_size(32.0)
                    .color("#FFFFFF")
                    .bold()
                    .glow(GlowProps::new(12.0, "#00B4D8", 0.7))
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(5.0, 1.4, 4.5, 0.8)
                .valign(AlignV::Middle)
                .build(),
        );

        // Outline effects row
        s.add_text(
            "OUTLINE EFFECTS",
            TextOptionsBuilder::new()
                .bounds(0.5, 2.4, 9.0, 0.4)
                .font_size(12.0)
                .color("#8FAADC")
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Red Outline")
                    .font_size(36.0)
                    .color("#FFFFFF")
                    .bold()
                    .outline(TextOutlineProps {
                        color: "FF0000".to_string(),
                        size: 1.5,
                    })
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(0.5, 2.8, 4.0, 0.8)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Gold Outline")
                    .font_size(36.0)
                    .color("#1B2A4A")
                    .bold()
                    .outline(TextOutlineProps {
                        color: "FFC000".to_string(),
                        size: 2.0,
                    })
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(5.0, 2.8, 4.5, 0.8)
                .valign(AlignV::Middle)
                .build(),
        );

        // Highlight effects row
        s.add_text(
            "HIGHLIGHT EFFECTS",
            TextOptionsBuilder::new()
                .bounds(0.5, 3.8, 9.0, 0.4)
                .font_size(12.0)
                .color("#8FAADC")
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Yellow ")
                    .font_size(20.0)
                    .color("#000000")
                    .highlight("#FFFF00")
                    .build(),
                TextRunBuilder::new("Cyan ")
                    .font_size(20.0)
                    .color("#000000")
                    .highlight("#00FFFF")
                    .build(),
                TextRunBuilder::new("Lime ")
                    .font_size(20.0)
                    .color("#000000")
                    .highlight("#00FF00")
                    .build(),
                TextRunBuilder::new("Pink ")
                    .font_size(20.0)
                    .color("#FFFFFF")
                    .highlight("#FF1493")
                    .build(),
                TextRunBuilder::new("Orange")
                    .font_size(20.0)
                    .color("#000000")
                    .highlight("#FF8C00")
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(0.5, 4.2, 9.0, 0.7)
                .valign(AlignV::Middle)
                .build(),
        );

        // Combined effects
        s.add_text_runs(
            vec![
                TextRunBuilder::new("Glow + Outline + Highlight")
                    .font_size(28.0)
                    .color("#FFFFFF")
                    .bold()
                    .glow(GlowProps::new(8.0, "#FF6B6B", 0.6))
                    .outline(TextOutlineProps {
                        color: "FF6B6B".to_string(),
                        size: 1.0,
                    })
                    .highlight("#2C1654")
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(0.5, 4.9, 9.0, 0.6)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Superscript, subscript, strikethrough, char spacing
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#FFFFFF");

        s.add_text(
            "Superscript, Subscript, Strikethrough & Spacing",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(24.0)
                .bold()
                .color("#1B2A4A")
                .build(),
        );

        // Chemical formulas with subscript
        s.add_text(
            "Chemical Formulas",
            TextOptionsBuilder::new()
                .bounds(0.5, 1.0, 4.0, 0.35)
                .font_size(13.0)
                .bold()
                .color("#4472C4")
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("H").font_size(24.0).color("#333333").build(),
                TextRunBuilder::new("2").font_size(24.0).color("#333333").subscript().build(),
                TextRunBuilder::new("O    C").font_size(24.0).color("#333333").build(),
                TextRunBuilder::new("6").font_size(24.0).color("#333333").subscript().build(),
                TextRunBuilder::new("H").font_size(24.0).color("#333333").build(),
                TextRunBuilder::new("12").font_size(24.0).color("#333333").subscript().build(),
                TextRunBuilder::new("O").font_size(24.0).color("#333333").build(),
                TextRunBuilder::new("6").font_size(24.0).color("#333333").subscript().build(),
                TextRunBuilder::new("    NaHCO").font_size(24.0).color("#333333").build(),
                TextRunBuilder::new("3").font_size(24.0).color("#333333").subscript().build(),
            ],
            TextOptionsBuilder::new()
                .bounds(0.5, 1.4, 4.5, 0.6)
                .valign(AlignV::Middle)
                .build(),
        );

        // Mathematical expressions with superscript
        s.add_text(
            "Math Expressions",
            TextOptionsBuilder::new()
                .bounds(5.5, 1.0, 4.0, 0.35)
                .font_size(13.0)
                .bold()
                .color("#4472C4")
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("E=mc").font_size(24.0).color("#333333").build(),
                TextRunBuilder::new("2").font_size(24.0).color("#333333").superscript().build(),
                TextRunBuilder::new("    x").font_size(24.0).color("#333333").build(),
                TextRunBuilder::new("2").font_size(24.0).color("#333333").superscript().build(),
                TextRunBuilder::new("+y").font_size(24.0).color("#333333").build(),
                TextRunBuilder::new("2").font_size(24.0).color("#333333").superscript().build(),
                TextRunBuilder::new("=r").font_size(24.0).color("#333333").build(),
                TextRunBuilder::new("2").font_size(24.0).color("#333333").superscript().build(),
            ],
            TextOptionsBuilder::new()
                .bounds(5.5, 1.4, 4.0, 0.6)
                .valign(AlignV::Middle)
                .build(),
        );

        // Strikethrough styles
        s.add_text(
            "Strikethrough",
            TextOptionsBuilder::new()
                .bounds(0.5, 2.3, 4.0, 0.35)
                .font_size(13.0)
                .bold()
                .color("#4472C4")
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Single strike")
                    .font_size(22.0)
                    .color("#CC0000")
                    .strike()
                    .build(),
                TextRunBuilder::new("    ")
                    .font_size(22.0)
                    .build(),
                TextRunBuilder::new("Double strike")
                    .font_size(22.0)
                    .color("#CC0000")
                    .strike_double()
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(0.5, 2.7, 9.0, 0.6)
                .valign(AlignV::Middle)
                .build(),
        );

        // Character spacing
        s.add_text(
            "Character Spacing",
            TextOptionsBuilder::new()
                .bounds(0.5, 3.5, 9.0, 0.35)
                .font_size(13.0)
                .bold()
                .color("#4472C4")
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Tight (-2pt)")
                    .font_size(18.0)
                    .color("#333333")
                    .char_spacing(-2.0)
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(0.5, 3.9, 3.0, 0.5)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Normal (0pt)")
                    .font_size(18.0)
                    .color("#333333")
                    .char_spacing(0.0)
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(3.5, 3.9, 3.0, 0.5)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Wide (+5pt)")
                    .font_size(18.0)
                    .color("#333333")
                    .char_spacing(5.0)
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(6.5, 3.9, 3.0, 0.5)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("VERY WIDE (+12pt)")
                    .font_size(18.0)
                    .color("#ED7D31")
                    .bold()
                    .char_spacing(12.0)
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(0.5, 4.5, 9.0, 0.6)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Text fit modes, vertical text, RTL
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#F5F6FA");

        s.add_text(
            "Text Fit, Vertical Direction & RTL",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color("#1B2A4A")
                .build(),
        );

        // Shrink to fit
        s.add_text(
            "SHRINK TO FIT",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.75, 2.5, 0.3)
                .font_size(10.0)
                .bold()
                .color("#4472C4")
                .build(),
        );

        s.add_shape(
            deckmint::ShapeType::RoundRect,
            deckmint::objects::shape::ShapeOptionsBuilder::new()
                .bounds(0.5, 1.1, 2.5, 1.4)
                .fill_color("#FFFFFF")
                .line_color("#4472C4")
                .line_width(1.5)
                .rect_radius(0.06)
                .build(),
        );
        s.add_text(
            "This is a long paragraph that demonstrates the shrink-to-fit text mode. \
             The font size will automatically reduce so all the text fits within this small box.",
            TextOptionsBuilder::new()
                .bounds(0.6, 1.15, 2.3, 1.3)
                .font_size(18.0)
                .color("#333333")
                .shrink_text()
                .build(),
        );

        // Auto-resize box
        s.add_text(
            "AUTO RESIZE",
            TextOptionsBuilder::new()
                .bounds(3.5, 0.75, 2.5, 0.3)
                .font_size(10.0)
                .bold()
                .color("#70AD47")
                .build(),
        );

        s.add_shape(
            deckmint::ShapeType::RoundRect,
            deckmint::objects::shape::ShapeOptionsBuilder::new()
                .bounds(3.5, 1.1, 2.5, 1.4)
                .fill_color("#FFFFFF")
                .line_color("#70AD47")
                .line_width(1.5)
                .rect_radius(0.06)
                .build(),
        );
        s.add_text(
            "Auto-resize adjusts the box height to fit this text.",
            TextOptionsBuilder::new()
                .bounds(3.6, 1.15, 2.3, 1.3)
                .font_size(14.0)
                .color("#333333")
                .autofit()
                .build(),
        );

        // Vertical text
        s.add_text(
            "VERTICAL TEXT",
            TextOptionsBuilder::new()
                .bounds(6.8, 0.75, 2.5, 0.3)
                .font_size(10.0)
                .bold()
                .color("#ED7D31")
                .build(),
        );

        s.add_shape(
            deckmint::ShapeType::RoundRect,
            deckmint::objects::shape::ShapeOptionsBuilder::new()
                .bounds(7.0, 1.1, 1.2, 3.5)
                .fill_color("#ED7D31")
                .rect_radius(0.06)
                .build(),
        );
        s.add_text(
            "Vertical Text",
            TextOptionsBuilder::new()
                .bounds(7.0, 1.1, 1.2, 3.5)
                .font_size(22.0)
                .bold()
                .color("#FFFFFF")
                .text_direction("vert")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_shape(
            deckmint::ShapeType::RoundRect,
            deckmint::objects::shape::ShapeOptionsBuilder::new()
                .bounds(8.5, 1.1, 1.2, 3.5)
                .fill_color("#5B9BD5")
                .rect_radius(0.06)
                .build(),
        );
        s.add_text(
            "Vert 270",
            TextOptionsBuilder::new()
                .bounds(8.5, 1.1, 1.2, 3.5)
                .font_size(22.0)
                .bold()
                .color("#FFFFFF")
                .text_direction("vert270")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // RTL text
        s.add_text(
            "RIGHT-TO-LEFT",
            TextOptionsBuilder::new()
                .bounds(0.5, 2.9, 4.0, 0.3)
                .font_size(10.0)
                .bold()
                .color("#9B59B6")
                .build(),
        );

        s.add_shape(
            deckmint::ShapeType::RoundRect,
            deckmint::objects::shape::ShapeOptionsBuilder::new()
                .bounds(0.5, 3.3, 5.5, 0.9)
                .fill_color("#FFFFFF")
                .line_color("#9B59B6")
                .line_width(1.5)
                .rect_radius(0.06)
                .build(),
        );
        s.add_text(
            "This text uses RTL (right-to-left) direction mode, useful for Arabic and Hebrew scripts.",
            TextOptionsBuilder::new()
                .bounds(0.6, 3.35, 5.3, 0.8)
                .font_size(14.0)
                .color("#333333")
                .rtl()
                .align(AlignH::Right)
                .valign(AlignV::Middle)
                .build(),
        );

        // Normal fit comparison
        s.add_text(
            "NORMAL (OVERFLOW)",
            TextOptionsBuilder::new()
                .bounds(0.5, 4.5, 2.5, 0.3)
                .font_size(10.0)
                .bold()
                .color("#CC0000")
                .build(),
        );

        s.add_shape(
            deckmint::ShapeType::RoundRect,
            deckmint::objects::shape::ShapeOptionsBuilder::new()
                .bounds(0.5, 4.85, 2.5, 0.65)
                .fill_color("#FFFFFF")
                .line_color("#CC0000")
                .line_width(1.5)
                .rect_radius(0.06)
                .build(),
        );
        s.add_text(
            "No text fit: this text may overflow beyond the bounds of this small box.",
            TextOptionsBuilder::new()
                .bounds(0.6, 4.9, 2.3, 0.55)
                .font_size(14.0)
                .color("#333333")
                .fit(TextFit::None)
                .build(),
        );
    }

    pres.write_to_file("25_text_effects.pptx").unwrap();
    println!("Wrote 25_text_effects.pptx");
}
