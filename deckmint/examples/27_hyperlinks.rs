//! Hyperlinks and navigation: URL links, slide-jump links, and navigation
//! actions. Creates an interactive-style presentation with a menu slide
//! and content slides with "back to menu" buttons.

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::{TextOptionsBuilder, TextRunBuilder};
use deckmint::types::{HyperlinkAction, HyperlinkProps};
use deckmint::{AlignH, AlignV, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Hyperlinks & Navigation".to_string();

    // ══════════════════════════════════════════════════════════
    // Slide 1: Title / Navigation menu
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#0D1B2A");

        // Accent bars
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.0, 10.0, 0.06)
                .fill_color("#00B4D8")
                .build(),
        );
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 5.565, 10.0, 0.06)
                .fill_color("#00B4D8")
                .build(),
        );

        s.add_text(
            "Interactive Presentation",
            TextOptionsBuilder::new()
                .bounds(1.0, 0.3, 8.0, 0.8)
                .font_size(36.0)
                .bold()
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_text(
            "Click any card below to navigate",
            TextOptionsBuilder::new()
                .bounds(1.0, 1.1, 8.0, 0.5)
                .font_size(16.0)
                .color("#90E0EF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Navigation cards to slides 2, 3, 4
        let nav_items = [
            ("About Us", "Learn about our team", "#4472C4", 2u32),
            ("Our Products", "Explore what we offer", "#ED7D31", 3),
            ("Resources", "External links & tools", "#70AD47", 4),
        ];

        let card_w = 2.6;
        let card_h = 2.3;
        let gap = 0.3;
        let total_w = 3.0 * card_w + 2.0 * gap;
        let x_start = (10.0 - total_w) / 2.0;
        let y_pos = 1.9;

        for (i, (title, desc, color, slide_num)) in nav_items.iter().enumerate() {
            let x = x_start + i as f64 * (card_w + gap);

            // Card background with hyperlink to slide
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(x, y_pos, card_w, card_h)
                    .fill_color("#1B2838")
                    .line_color(*color)
                    .line_width(2.0)
                    .rect_radius(0.12)
                    .hyperlink(
                        HyperlinkProps::slide(*slide_num)
                            .with_tooltip(&format!("Go to {}", title)),
                    )
                    .build(),
            );

            // Number circle
            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(x + (card_w - 0.6) / 2.0, y_pos + 0.3, 0.6, 0.6)
                    .fill_color(*color)
                    .build(),
            );
            s.add_text(
                &format!("{}", i + 1),
                TextOptionsBuilder::new()
                    .bounds(x + (card_w - 0.6) / 2.0, y_pos + 0.3, 0.6, 0.6)
                    .font_size(22.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Title
            s.add_text(
                *title,
                TextOptionsBuilder::new()
                    .bounds(x + 0.15, y_pos + 1.1, card_w - 0.3, 0.5)
                    .font_size(18.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Description
            s.add_text(
                *desc,
                TextOptionsBuilder::new()
                    .bounds(x + 0.15, y_pos + 1.6, card_w - 0.3, 0.5)
                    .font_size(12.0)
                    .color("#A0B4C8")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }

        // Navigation hint at bottom
        s.add_text_runs(
            vec![
                TextRunBuilder::new("Tip: use ")
                    .font_size(11.0)
                    .color("#607080")
                    .build(),
                TextRunBuilder::new("Next/Prev")
                    .font_size(11.0)
                    .color("#90E0EF")
                    .bold()
                    .build(),
                TextRunBuilder::new(" buttons on content slides to navigate sequentially")
                    .font_size(11.0)
                    .color("#607080")
                    .build(),
            ],
            TextOptionsBuilder::new()
                .bounds(1.0, 4.8, 8.0, 0.5)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: About Us (with back-to-menu link)
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#FFFFFF");

        // Header bar
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.0, 10.0, 0.8)
                .fill_color("#4472C4")
                .build(),
        );
        s.add_text(
            "About Us",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.0, 5.0, 0.8)
                .font_size(28.0)
                .bold()
                .color("#FFFFFF")
                .valign(AlignV::Middle)
                .build(),
        );

        // Back to menu button
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(7.5, 0.15, 2.2, 0.5)
                .fill_color("#2C5F8A")
                .rect_radius(0.06)
                .hyperlink(
                    HyperlinkProps::slide(1).with_tooltip("Return to main menu"),
                )
                .build(),
        );
        s.add_text(
            "< Back to Menu",
            TextOptionsBuilder::new()
                .bounds(7.5, 0.15, 2.2, 0.5)
                .font_size(12.0)
                .bold()
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Content
        s.add_text(
            "We are a team of passionate developers building tools that make \
             presentations easy to create programmatically.",
            TextOptionsBuilder::new()
                .bounds(0.8, 1.2, 8.4, 1.0)
                .font_size(18.0)
                .color("#333333")
                .build(),
        );

        // Team members with slide links
        let team = [
            ("Alice Johnson", "Lead Developer", "#4472C4"),
            ("Bob Smith", "Design Engineer", "#ED7D31"),
            ("Carol Lee", "Product Manager", "#70AD47"),
        ];

        for (i, (name, role, color)) in team.iter().enumerate() {
            let x = 0.8 + i as f64 * 3.0;
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(x, 2.5, 2.6, 1.4)
                    .fill_color(*color)
                    .rect_radius(0.08)
                    .build(),
            );
            s.add_text(
                *name,
                TextOptionsBuilder::new()
                    .bounds(x + 0.15, 2.7, 2.3, 0.5)
                    .font_size(16.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
            s.add_text(
                *role,
                TextOptionsBuilder::new()
                    .bounds(x + 0.15, 3.2, 2.3, 0.5)
                    .font_size(12.0)
                    .color("#E0E8F0")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }

        // Nav buttons at bottom
        add_nav_bar(s, true);
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Products (with external URL links)
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#FFFFFF");

        // Header bar
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.0, 10.0, 0.8)
                .fill_color("#ED7D31")
                .build(),
        );
        s.add_text(
            "Our Products",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.0, 5.0, 0.8)
                .font_size(28.0)
                .bold()
                .color("#FFFFFF")
                .valign(AlignV::Middle)
                .build(),
        );

        // Back to menu button
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(7.5, 0.15, 2.2, 0.5)
                .fill_color("#C06020")
                .rect_radius(0.06)
                .hyperlink(
                    HyperlinkProps::slide(1).with_tooltip("Return to main menu"),
                )
                .build(),
        );
        s.add_text(
            "< Back to Menu",
            TextOptionsBuilder::new()
                .bounds(7.5, 0.15, 2.2, 0.5)
                .font_size(12.0)
                .bold()
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Product cards with external URL hyperlinks on text
        s.add_text(
            "Click any product name to visit its website:",
            TextOptionsBuilder::new()
                .bounds(0.8, 1.1, 8.4, 0.5)
                .font_size(16.0)
                .color("#555555")
                .build(),
        );

        let products = [
            ("deckmint", "Rust PowerPoint generator", "https://github.com/nickabal/pptxgen-rs", "#4472C4"),
            ("Rust Lang", "Systems programming language", "https://www.rust-lang.org", "#ED7D31"),
            ("GitHub", "Code hosting platform", "https://github.com", "#333333"),
        ];

        for (i, (name, desc, url, color)) in products.iter().enumerate() {
            let y = 1.8 + i as f64 * 1.1;

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(0.8, y, 8.4, 0.9)
                    .fill_color("#F8F9FA")
                    .line_color(*color)
                    .line_width(1.5)
                    .rect_radius(0.06)
                    .build(),
            );

            // Product name as clickable hyperlink
            s.add_text_runs(
                vec![
                    TextRunBuilder::new(*name)
                        .font_size(18.0)
                        .bold()
                        .color(*color)
                        .hyperlink(HyperlinkProps::url(*url).with_tooltip(*url))
                        .build(),
                    TextRunBuilder::new(&format!("  -  {}", desc))
                        .font_size(14.0)
                        .color("#666666")
                        .build(),
                ],
                TextOptionsBuilder::new()
                    .bounds(1.1, y + 0.1, 7.8, 0.35)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // URL display
            s.add_text_runs(
                vec![TextRunBuilder::new(*url)
                    .font_size(11.0)
                    .color("#999999")
                    .hyperlink(HyperlinkProps::url(*url))
                    .build()],
                TextOptionsBuilder::new()
                    .bounds(1.1, y + 0.45, 7.8, 0.35)
                    .valign(AlignV::Top)
                    .build(),
            );
        }

        add_nav_bar(s, true);
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Resources (navigation actions)
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#FFFFFF");

        // Header bar
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.0, 10.0, 0.8)
                .fill_color("#70AD47")
                .build(),
        );
        s.add_text(
            "Resources & Navigation Actions",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.0, 7.0, 0.8)
                .font_size(24.0)
                .bold()
                .color("#FFFFFF")
                .valign(AlignV::Middle)
                .build(),
        );

        // Back to menu button
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(7.5, 0.15, 2.2, 0.5)
                .fill_color("#4A8A2C")
                .rect_radius(0.06)
                .hyperlink(
                    HyperlinkProps::slide(1).with_tooltip("Return to main menu"),
                )
                .build(),
        );
        s.add_text(
            "< Back to Menu",
            TextOptionsBuilder::new()
                .bounds(7.5, 0.15, 2.2, 0.5)
                .font_size(12.0)
                .bold()
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Section: Navigation Actions
        s.add_text(
            "Navigation Action Buttons",
            TextOptionsBuilder::new()
                .bounds(0.8, 1.0, 8.4, 0.5)
                .font_size(18.0)
                .bold()
                .color("#333333")
                .build(),
        );

        s.add_text(
            "These buttons use built-in PowerPoint navigation actions:",
            TextOptionsBuilder::new()
                .bounds(0.8, 1.45, 8.4, 0.4)
                .font_size(13.0)
                .color("#666666")
                .build(),
        );

        let actions: Vec<(&str, HyperlinkAction, &str)> = vec![
            ("First Slide", HyperlinkAction::FirstSlide, "#4472C4"),
            ("Prev Slide", HyperlinkAction::PrevSlide, "#ED7D31"),
            ("Next Slide", HyperlinkAction::NextSlide, "#70AD47"),
            ("Last Slide", HyperlinkAction::LastSlide, "#9B59B6"),
            ("End Show", HyperlinkAction::EndShow, "#E74C3C"),
        ];

        let btn_w = 1.6;
        let btn_gap = 0.2;
        let total_btn_w = actions.len() as f64 * btn_w + (actions.len() as f64 - 1.0) * btn_gap;
        let btn_x_start = (10.0 - total_btn_w) / 2.0;

        for (i, (label, action, color)) in actions.iter().enumerate() {
            let x = btn_x_start + i as f64 * (btn_w + btn_gap);

            let mut hl = HyperlinkProps::url("");
            hl.action = Some(action.clone());

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(x, 2.0, btn_w, 0.55)
                    .fill_color(*color)
                    .rect_radius(0.06)
                    .hyperlink(hl)
                    .build(),
            );
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(x, 2.0, btn_w, 0.55)
                    .font_size(11.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // External resources with text hyperlinks
        s.add_text(
            "Useful External Links",
            TextOptionsBuilder::new()
                .bounds(0.8, 3.0, 8.4, 0.5)
                .font_size(18.0)
                .bold()
                .color("#333333")
                .build(),
        );

        let links = [
            ("Rust Documentation", "https://doc.rust-lang.org"),
            ("OOXML Specification", "https://www.ecma-international.org"),
            ("PowerPoint File Format", "https://learn.microsoft.com"),
        ];

        for (i, (text, url)) in links.iter().enumerate() {
            let y = 3.5 + i as f64 * 0.55;
            s.add_text_runs(
                vec![
                    TextRunBuilder::new("\u{2192}  ")
                        .font_size(15.0)
                        .color("#70AD47")
                        .build(),
                    TextRunBuilder::new(*text)
                        .font_size(15.0)
                        .color("#4472C4")
                        .underline("sng")
                        .hyperlink(HyperlinkProps::url(*url).with_tooltip(*url))
                        .build(),
                ],
                TextOptionsBuilder::new()
                    .bounds(1.2, y, 7.6, 0.45)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }
    }

    pres.write_to_file("27_hyperlinks.pptx").unwrap();
    println!("Wrote 27_hyperlinks.pptx");
}

/// Add a bottom navigation bar with Prev/Next action buttons.
fn add_nav_bar(s: &mut deckmint::Slide, show_both: bool) {
    let bar_y = 5.05;

    // Separator line
    s.add_shape(
        ShapeType::Rect,
        ShapeOptionsBuilder::new()
            .bounds(0.5, bar_y - 0.08, 9.0, 0.02)
            .fill_color("#E0E0E0")
            .build(),
    );

    if show_both {
        // Previous button
        let mut prev_hl = HyperlinkProps::url("");
        prev_hl.action = Some(HyperlinkAction::PrevSlide);

        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(0.5, bar_y, 1.6, 0.45)
                .fill_color("#E8ECF1")
                .rect_radius(0.06)
                .hyperlink(prev_hl)
                .build(),
        );
        s.add_text(
            "\u{2190} Previous",
            TextOptionsBuilder::new()
                .bounds(0.5, bar_y, 1.6, 0.45)
                .font_size(11.0)
                .bold()
                .color("#555555")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Next button
        let mut next_hl = HyperlinkProps::url("");
        next_hl.action = Some(HyperlinkAction::NextSlide);

        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(7.9, bar_y, 1.6, 0.45)
                .fill_color("#E8ECF1")
                .rect_radius(0.06)
                .hyperlink(next_hl)
                .build(),
        );
        s.add_text(
            "Next \u{2192}",
            TextOptionsBuilder::new()
                .bounds(7.9, bar_y, 1.6, 0.45)
                .font_size(11.0)
                .bold()
                .color("#555555")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }
}
