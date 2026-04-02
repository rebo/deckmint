//! Visually stunning infographic-style presentation using grid layouts,
//! shapes, and text. Includes a dark-themed title, timeline/process flow,
//! statistics dashboard, and comparison layout.

use deckmint::layout::GridLayoutBuilder;
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, GradientFill, Presentation, ShapeType, SlideMasterDef};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Infographic Presentation".to_string();

    // ── Master slide for consistent branding ────────────────
    {
        let mut master = SlideMasterDef::new("Infographic");
        master.set_background_color("#0F1923");

        // Top accent line
        master.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.0, 10.0, 0.04)
                .fill_color("#00D2FF")
                .build(),
        );

        // Bottom accent line
        master.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 5.585, 10.0, 0.04)
                .fill_color("#00D2FF")
                .build(),
        );

        // Side accent
        master.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.04, 0.03, 5.545)
                .fill_color("#7B2FBE")
                .build(),
        );

        pres.define_master(master);
    }

    // ══════════════════════════════════════════════════════════
    // Slide 1: Title slide with dark background and accent colors
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        // Large gradient accent shape
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 1.6, 10.0, 2.4)
                .gradient_fill(GradientFill::two_color(0.0, "#7B2FBE", "#00D2FF"))
                .build(),
        );

        // Dark overlay for text readability
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 1.6, 10.0, 2.4)
                .fill_color("#0F1923")
                .build(),
        );

        // Re-add a lighter gradient strip
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.5, 2.0, 9.0, 0.04)
                .gradient_fill(GradientFill::two_color(0.0, "#7B2FBE", "#00D2FF"))
                .build(),
        );

        s.add_text(
            "THE STATE OF TECHNOLOGY",
            TextOptionsBuilder::new()
                .bounds(1.0, 0.5, 8.0, 0.8)
                .font_size(14.0)
                .bold()
                .color("#00D2FF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_text(
            "2025 Annual Report",
            TextOptionsBuilder::new()
                .bounds(1.0, 2.2, 8.0, 1.2)
                .font_size(44.0)
                .bold()
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(3.5, 3.6, 3.0, 0.04)
                .gradient_fill(GradientFill::two_color(0.0, "#7B2FBE", "#00D2FF"))
                .build(),
        );

        s.add_text(
            "Insights, Trends & Key Metrics",
            TextOptionsBuilder::new()
                .bounds(1.0, 3.8, 8.0, 0.6)
                .font_size(18.0)
                .color("#8FAADC")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Decorative elements at bottom
        let metrics = [
            ("150+", "Countries"),
            ("$2.4B", "Revenue"),
            ("98%", "Uptime"),
            ("10M+", "Users"),
        ];

        let metric_w = 2.0;
        let metric_gap = 0.2;
        let total = metrics.len() as f64 * metric_w + (metrics.len() as f64 - 1.0) * metric_gap;
        let mx_start = (10.0 - total) / 2.0;

        for (i, (value, label)) in metrics.iter().enumerate() {
            let x = mx_start + i as f64 * (metric_w + metric_gap);
            s.add_text(
                *value,
                TextOptionsBuilder::new()
                    .bounds(x, 4.6, metric_w, 0.5)
                    .font_size(22.0)
                    .bold()
                    .color("#00D2FF")
                    .align(AlignH::Center)
                    .valign(AlignV::Bottom)
                    .build(),
            );
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(x, 5.05, metric_w, 0.35)
                    .font_size(11.0)
                    .color("#607080")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Timeline / Process flow
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "OUR JOURNEY",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.5)
                .font_size(12.0)
                .bold()
                .color("#00D2FF")
                .build(),
        );

        s.add_text(
            "Key Milestones",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.55, 9.0, 0.5)
                .font_size(28.0)
                .bold()
                .color("#FFFFFF")
                .build(),
        );

        // Horizontal connector line
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(1.0, 2.35, 8.0, 0.04)
                .gradient_fill(GradientFill::two_color(0.0, "#7B2FBE", "#00D2FF"))
                .build(),
        );

        let milestones = [
            ("2020", "Founded", "Started with a\nsmall team of 5", "#7B2FBE"),
            ("2021", "Series A", "Raised $12M in\nfirst funding round", "#9B44E0"),
            ("2022", "Global", "Expanded to\n50+ countries", "#4A90D9"),
            ("2023", "Scale", "Reached 1M+\nactive users", "#00B4D8"),
            ("2024", "IPO", "Went public on\nNASDAQ exchange", "#00D2FF"),
        ];

        let step_w = 1.6;
        let step_gap = 0.2;
        let total_w = milestones.len() as f64 * step_w + (milestones.len() as f64 - 1.0) * step_gap;
        let sx_start = (10.0 - total_w) / 2.0;

        for (i, (year, title, desc, color)) in milestones.iter().enumerate() {
            let cx = sx_start + i as f64 * (step_w + step_gap) + step_w / 2.0;

            // Circle node on the timeline
            let circle_r = 0.25;
            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(cx - circle_r, 2.37 - circle_r, circle_r * 2.0, circle_r * 2.0)
                    .fill_color(*color)
                    .build(),
            );

            // Step number inside circle
            s.add_text(
                &format!("{}", i + 1),
                TextOptionsBuilder::new()
                    .bounds(cx - circle_r, 2.37 - circle_r, circle_r * 2.0, circle_r * 2.0)
                    .font_size(14.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Year above
            s.add_text(
                *year,
                TextOptionsBuilder::new()
                    .bounds(cx - step_w / 2.0, 1.5, step_w, 0.5)
                    .font_size(18.0)
                    .bold()
                    .color(*color)
                    .align(AlignH::Center)
                    .valign(AlignV::Bottom)
                    .build(),
            );

            // Title below
            s.add_text(
                *title,
                TextOptionsBuilder::new()
                    .bounds(cx - step_w / 2.0, 2.9, step_w, 0.4)
                    .font_size(14.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );

            // Description
            s.add_text(
                *desc,
                TextOptionsBuilder::new()
                    .bounds(cx - step_w / 2.0, 3.3, step_w, 0.8)
                    .font_size(10.0)
                    .color("#8FAADC")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }

        // Bottom decorative bar
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.5, 4.6, 9.0, 0.6)
                .fill_color("#141E2A")
                .build(),
        );
        s.add_text(
            "5 years of continuous growth and innovation",
            TextOptionsBuilder::new()
                .bounds(0.5, 4.6, 9.0, 0.6)
                .font_size(14.0)
                .italic()
                .color("#607080")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Statistics dashboard
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "BY THE NUMBERS",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color("#00D2FF")
                .build(),
        );

        s.add_text(
            "Key Performance Metrics",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.5, 9.0, 0.5)
                .font_size(26.0)
                .bold()
                .color("#FFFFFF")
                .build(),
        );

        // Large stat cards - top row
        let top_grid = GridLayoutBuilder::grid_n_m(3, 1, 0.2)
            .origin(0.5, 1.2)
            .container(9.0, 1.8)
            .build();

        let top_stats = [
            ("$2.4B", "Total Revenue", "+34% YoY", "#7B2FBE"),
            ("10.2M", "Active Users", "+67% YoY", "#00B4D8"),
            ("99.97%", "Platform Uptime", "Industry Leading", "#00D2FF"),
        ];

        for (i, (value, label, growth, color)) in top_stats.iter().enumerate() {
            let cell = top_grid.cell(i, 0);

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color("#141E2A")
                    .line_color(*color)
                    .line_width(1.5)
                    .rect_radius(0.08)
                    .build(),
            );

            // Accent bar at top
            let accent = deckmint::layout::CellRect {
                x: cell.x,
                y: cell.y,
                w: cell.w,
                h: 0.05,
            };
            s.add_shape(
                ShapeType::Rect,
                ShapeOptionsBuilder::new()
                    .rect(accent)
                    .fill_color(*color)
                    .build(),
            );

            let inner = cell.inset(0.15);
            s.add_text(
                *value,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y, inner.w, 0.7)
                    .font_size(36.0)
                    .bold()
                    .color(*color)
                    .align(AlignH::Center)
                    .valign(AlignV::Bottom)
                    .build(),
            );
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y + 0.7, inner.w, 0.35)
                    .font_size(13.0)
                    .color("#8FAADC")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
            s.add_text(
                *growth,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y + 1.0, inner.w, 0.35)
                    .font_size(11.0)
                    .bold()
                    .color("#70AD47")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }

        // Bottom row: smaller icon-like stat boxes
        let bot_grid = GridLayoutBuilder::grid_n_m(4, 1, 0.15)
            .origin(0.5, 3.3)
            .container(9.0, 2.0)
            .build();

        let bot_stats = [
            ("\u{2605}", "4.8/5", "App Rating", "#FFC000"),
            ("\u{2191}", "156%", "API Growth", "#00D2FF"),
            ("\u{2764}", "94%", "Satisfaction", "#E74C3C"),
            ("\u{26A1}", "< 50ms", "Response Time", "#70AD47"),
        ];

        for (i, (icon, value, label, color)) in bot_stats.iter().enumerate() {
            let cell = bot_grid.cell(i, 0);

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color("#141E2A")
                    .rect_radius(0.08)
                    .build(),
            );

            // Icon shape (circle background)
            let icon_size = 0.5;
            let icon_x = cell.x + (cell.w - icon_size) / 2.0;
            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(icon_x, cell.y + 0.2, icon_size, icon_size)
                    .fill_color(*color)
                    .build(),
            );
            s.add_text(
                *icon,
                TextOptionsBuilder::new()
                    .bounds(icon_x, cell.y + 0.2, icon_size, icon_size)
                    .font_size(18.0)
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            s.add_text(
                *value,
                TextOptionsBuilder::new()
                    .bounds(cell.x, cell.y + 0.85, cell.w, 0.5)
                    .font_size(24.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Bottom)
                    .build(),
            );
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(cell.x, cell.y + 1.35, cell.w, 0.4)
                    .font_size(11.0)
                    .color("#607080")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Comparison layout (vs. style)
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "COMPARISON",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color("#00D2FF")
                .build(),
        );

        s.add_text(
            "Before vs. After",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.45, 9.0, 0.5)
                .font_size(28.0)
                .bold()
                .color("#FFFFFF")
                .build(),
        );

        // VS circle in the center
        s.add_shape(
            ShapeType::Ellipse,
            ShapeOptionsBuilder::new()
                .bounds(4.5, 2.3, 1.0, 1.0)
                .fill_color("#0F1923")
                .line_color("#00D2FF")
                .line_width(2.5)
                .build(),
        );
        s.add_text(
            "VS",
            TextOptionsBuilder::new()
                .bounds(4.5, 2.3, 1.0, 1.0)
                .font_size(22.0)
                .bold()
                .color("#00D2FF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Left panel: "Before"
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(0.4, 1.2, 4.0, 3.8)
                .fill_color("#1A1520")
                .line_color("#7B2FBE")
                .line_width(1.5)
                .rect_radius(0.1)
                .build(),
        );

        s.add_text(
            "BEFORE",
            TextOptionsBuilder::new()
                .bounds(0.4, 1.3, 4.0, 0.5)
                .font_size(20.0)
                .bold()
                .color("#7B2FBE")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        let before_items = [
            ("Manual process", "#E74C3C"),
            ("3 days turnaround", "#E74C3C"),
            ("High error rate", "#E74C3C"),
            ("Limited scale", "#E74C3C"),
        ];

        for (i, (text, color)) in before_items.iter().enumerate() {
            let y = 2.0 + i as f64 * 0.6;
            // X mark
            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(0.8, y, 0.3, 0.3)
                    .fill_color(*color)
                    .build(),
            );
            s.add_text(
                "\u{2717}",
                TextOptionsBuilder::new()
                    .bounds(0.8, y, 0.3, 0.3)
                    .font_size(12.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
            s.add_text(
                *text,
                TextOptionsBuilder::new()
                    .bounds(1.25, y, 2.8, 0.3)
                    .font_size(14.0)
                    .color("#C0C0C0")
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Right panel: "After"
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(5.6, 1.2, 4.0, 3.8)
                .fill_color("#0A1A20")
                .line_color("#00D2FF")
                .line_width(1.5)
                .rect_radius(0.1)
                .build(),
        );

        s.add_text(
            "AFTER",
            TextOptionsBuilder::new()
                .bounds(5.6, 1.3, 4.0, 0.5)
                .font_size(20.0)
                .bold()
                .color("#00D2FF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        let after_items = [
            ("Fully automated", "#00D2FF"),
            ("Real-time delivery", "#00D2FF"),
            ("99.9% accuracy", "#00D2FF"),
            ("Unlimited scale", "#00D2FF"),
        ];

        for (i, (text, color)) in after_items.iter().enumerate() {
            let y = 2.0 + i as f64 * 0.6;
            // Check mark
            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(6.0, y, 0.3, 0.3)
                    .fill_color(*color)
                    .build(),
            );
            s.add_text(
                "\u{2713}",
                TextOptionsBuilder::new()
                    .bounds(6.0, y, 0.3, 0.3)
                    .font_size(12.0)
                    .bold()
                    .color("#0F1923")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
            s.add_text(
                *text,
                TextOptionsBuilder::new()
                    .bounds(6.45, y, 2.8, 0.3)
                    .font_size(14.0)
                    .color("#FFFFFF")
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Bottom summary bar
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(0.4, 5.1, 9.2, 0.4)
                .gradient_fill(GradientFill::two_color(0.0, "#7B2FBE", "#00D2FF"))
                .rect_radius(0.04)
                .build(),
        );
        s.add_text(
            "Result: 10x faster  |  85% cost reduction  |  Zero manual errors",
            TextOptionsBuilder::new()
                .bounds(0.4, 5.1, 9.2, 0.4)
                .font_size(13.0)
                .bold()
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    pres.write_to_file("28_infographic.pptx").unwrap();
    println!("Wrote 28_infographic.pptx");
}
