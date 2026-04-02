//! 5-slide creative portfolio — vibrant coral/purple/teal palette with gradient
//! accents, photo, progress bars, project cards, and shape-based contact icons.

use deckmint::layout::{split_h, split_v, GridLayoutBuilder};
use deckmint::objects::image::ImageOptionsBuilder;
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::{TextOptionsBuilder, TextRunBuilder};
use deckmint::{
    AlignH, AlignV, GradientFill, Presentation, ShapeType, SlideMasterDef,
};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Creative Portfolio".to_string();

    // ── Colors ────────────────────────────────────────────────
    let bg = "#FAFAFA";
    let dark = "#1A1A2E";
    let coral = "#E84855";
    let purple = "#7209B7";
    let teal = "#2EC4B6";
    let white = "#FFFFFF";
    let muted = "#666680";
    let light_gray = "#F0F0F5";

    // ── Load tiger.png ────────────────────────────────────────
    let photo = std::fs::read(concat!(env!("CARGO_MANIFEST_DIR"), "/examples/tiger.png"))
        .expect("tiger.png not found — place it in deckmint/examples/");

    // ── Master slide ──────────────────────────────────────────
    {
        let mut master = SlideMasterDef::new("Portfolio");
        master.set_background_color(bg);

        // Bottom accent strip
        master.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 5.525, 10.0, 0.1)
                .gradient_fill(GradientFill::two_color(0.0, coral, purple))
                .build(),
        );

        pres.define_master(master);
    }

    // ══════════════════════════════════════════════════════════
    // Slide 1: Title — bold name with gradient accent shapes
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        // Large gradient rectangle as decorative backdrop
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.0, 10.0, 3.2)
                .gradient_fill(GradientFill::two_color(45.0, purple, coral))
                .build(),
        );

        // Decorative teal accent shape (overlapping)
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(6.5, 2.4, 3.2, 1.6)
                .fill_color(teal)
                .rect_radius(0.1)
                .build(),
        );

        // Another accent triangle
        s.add_shape(
            ShapeType::Triangle,
            ShapeOptionsBuilder::new()
                .bounds(0.3, 2.0, 1.5, 1.5)
                .fill_color(teal)
                .build(),
        );

        // Name
        s.add_text(
            "ALEX MORGAN",
            TextOptionsBuilder::new()
                .bounds(0.8, 0.8, 8.4, 1.0)
                .font_size(52.0)
                .bold()
                .color(white)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Title line
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(3.5, 1.85, 3.0, 0.04)
                .fill_color(white)
                .build(),
        );

        s.add_text(
            "Creative Designer & Developer",
            TextOptionsBuilder::new()
                .bounds(1.0, 2.0, 8.0, 0.6)
                .font_size(20.0)
                .color(white)
                .align(AlignH::Center)
                .valign(AlignV::Top)
                .build(),
        );

        // Bottom tagline
        s.add_text(
            "Crafting beautiful digital experiences since 2018",
            TextOptionsBuilder::new()
                .bounds(1.0, 3.8, 8.0, 0.5)
                .font_size(16.0)
                .color(muted)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Social handles
        s.add_text(
            "alexmorgan.design  |  @alexmorgan  |  github.com/alexmorgan",
            TextOptionsBuilder::new()
                .bounds(1.0, 4.5, 8.0, 0.4)
                .font_size(11.0)
                .color(purple)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: About — photo + bio text
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "ABOUT ME",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(coral)
                .build(),
        );

        // Two-column layout
        let grid = GridLayoutBuilder::grid_n_m(2, 1, 0.4)
            .origin(0.5, 0.8)
            .container(9.0, 4.4)
            .build();

        let left = grid.cell(0, 0);
        let right = grid.cell(1, 0);

        // Photo with decorative frame
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(left.x + 0.2, left.y + 0.2, left.w - 0.4, 3.4)
                .fill_color(purple)
                .rect_radius(0.1)
                .build(),
        );

        // Tiger photo (offset slightly for frame effect)
        s.add_image_from(
            ImageOptionsBuilder::new()
                .bytes(photo.clone(), "png")
                .bounds(left.x + 0.3, left.y + 0.3, left.w - 0.6, 3.2)
                .rounding(),
        )
        .unwrap();

        // Decorative shape
        s.add_shape(
            ShapeType::Ellipse,
            ShapeOptionsBuilder::new()
                .bounds(left.x + left.w - 0.8, left.y + 3.2, 0.8, 0.8)
                .fill_color(teal)
                .build(),
        );

        // Bio text (right side)
        s.add_text(
            "Hello!",
            TextOptionsBuilder::new()
                .bounds(right.x, right.y + 0.1, right.w, 0.5)
                .font_size(28.0)
                .bold()
                .color(dark)
                .build(),
        );

        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(right.x, right.y + 0.65, 1.5, 0.04)
                .gradient_fill(GradientFill::two_color(0.0, coral, purple))
                .build(),
        );

        s.add_text(
            "I'm a creative designer and developer with 7+ years of experience crafting digital products that delight users.\n\nMy passion lies at the intersection of beautiful design and clean code. I've worked with startups and Fortune 500 companies to build products used by millions.\n\nWhen I'm not designing, you'll find me hiking, photographing wildlife, or contributing to open-source projects.",
            TextOptionsBuilder::new()
                .bounds(right.x, right.y + 0.9, right.w, 2.8)
                .font_size(11.5)
                .color(muted)
                .build(),
        );

        // Highlight stats
        let stats_area = deckmint::layout::CellRect {
            x: right.x,
            y: right.y + 3.5,
            w: right.w,
            h: 0.7,
        };
        let stat_cols = split_h(&stats_area, 3, 0.15);

        let stats = [
            ("50+", "Projects", coral),
            ("30+", "Clients", purple),
            ("7+", "Years", teal),
        ];

        for (i, (value, label, color)) in stats.iter().enumerate() {
            let cell = stat_cols[i];
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color(*color)
                    .rect_radius(0.04)
                    .build(),
            );
            s.add_text_runs(
                vec![
                    TextRunBuilder::new(*value)
                        .font_size(18.0)
                        .bold()
                        .color(white)
                        .build(),
                    TextRunBuilder::new(&format!(" {}", label))
                        .font_size(10.0)
                        .color(white)
                        .build(),
                ],
                TextOptionsBuilder::new()
                    .rect(cell)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Skills — progress bars (colored rect over gray rect)
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "SKILLS",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(coral)
                .build(),
        );

        s.add_text(
            "What I Bring to the Table",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.5, 9.0, 0.5)
                .font_size(26.0)
                .bold()
                .color(dark)
                .build(),
        );

        let skills = [
            ("UI/UX Design", 95.0, coral),
            ("Frontend Development", 90.0, purple),
            ("React / TypeScript", 88.0, teal),
            ("Brand Identity", 85.0, coral),
            ("Motion Design", 78.0, purple),
            ("Backend / APIs", 72.0, teal),
            ("3D / WebGL", 65.0, coral),
        ];

        let bar_area = deckmint::layout::CellRect {
            x: 0.8,
            y: 1.2,
            w: 8.4,
            h: 4.0,
        };
        let skill_rows = split_v(&bar_area, skills.len(), 0.1);

        for (i, (name, pct, color)) in skills.iter().enumerate() {
            let row = skill_rows[i];
            let label_w = 2.5;
            let bar_x = row.x + label_w;
            let bar_w = row.w - label_w - 0.8;
            let bar_h = row.h * 0.55;
            let bar_y = row.y + (row.h - bar_h) / 2.0;

            // Skill name
            s.add_text(
                *name,
                TextOptionsBuilder::new()
                    .bounds(row.x, row.y, label_w, row.h)
                    .font_size(12.0)
                    .bold()
                    .color(dark)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Gray background bar
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(bar_x, bar_y, bar_w, bar_h)
                    .fill_color(light_gray)
                    .rect_radius(bar_h / 2.0)
                    .build(),
            );

            // Filled progress bar
            let filled_w = bar_w * (pct / 100.0);
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(bar_x, bar_y, filled_w, bar_h)
                    .fill_color(*color)
                    .rect_radius(bar_h / 2.0)
                    .build(),
            );

            // Percentage label
            s.add_text(
                &format!("{}%", *pct as u32),
                TextOptionsBuilder::new()
                    .bounds(bar_x + bar_w + 0.1, row.y, 0.7, row.h)
                    .font_size(11.0)
                    .bold()
                    .color(*color)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Projects — 2x2 card grid
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "PROJECTS",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(coral)
                .build(),
        );

        s.add_text(
            "Selected Work",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.5, 9.0, 0.45)
                .font_size(24.0)
                .bold()
                .color(dark)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(2, 2, 0.2)
            .origin(0.5, 1.1)
            .container(9.0, 4.2)
            .build();

        let projects = [
            ("Lumina App", "Mobile banking app redesign for 2M+ users.\nClean UI, biometric login, real-time analytics.", coral, "Mobile Design"),
            ("NovaBrand", "Complete brand identity for a tech startup.\nLogo, guidelines, marketing collateral.", purple, "Branding"),
            ("DataViz Pro", "Interactive data visualization dashboard.\nD3.js, React, real-time WebSocket updates.", teal, "Development"),
            ("EcoTrack", "Sustainability tracking platform.\nGamified UX with social sharing features.", "#4CAF50", "Full Stack"),
        ];

        for (idx, (name, desc, color, tag)) in projects.iter().enumerate() {
            let col = idx % 2;
            let row = idx / 2;
            let cell = grid.cell(col, row);

            // Card background
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color(white)
                    .line_color("#E0E0E5")
                    .line_width(1.0)
                    .rect_radius(0.08)
                    .build(),
            );

            // Color accent bar at top
            s.add_shape(
                ShapeType::Rect,
                ShapeOptionsBuilder::new()
                    .bounds(cell.x, cell.y, cell.w, 0.06)
                    .fill_color(*color)
                    .build(),
            );

            let inner = cell.inset(0.2);

            // Tag pill
            let tag_w = tag.len() as f64 * 0.09 + 0.3;
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(inner.x, inner.y, tag_w, 0.28)
                    .fill_color(*color)
                    .rect_radius(0.14)
                    .build(),
            );
            s.add_text(
                *tag,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y, tag_w, 0.28)
                    .font_size(8.0)
                    .bold()
                    .color(white)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Project name
            s.add_text(
                *name,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y + 0.4, inner.w, 0.4)
                    .font_size(16.0)
                    .bold()
                    .color(dark)
                    .valign(AlignV::Bottom)
                    .build(),
            );

            // Separator line
            s.add_shape(
                ShapeType::Rect,
                ShapeOptionsBuilder::new()
                    .bounds(inner.x, inner.y + 0.85, 1.2, 0.03)
                    .fill_color(*color)
                    .build(),
            );

            // Description
            s.add_text(
                *desc,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y + 0.95, inner.w, 0.8)
                    .font_size(10.0)
                    .color(muted)
                    .valign(AlignV::Top)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 5: Contact — shape-based icons + info
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        // Gradient header area
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.0, 10.0, 2.2)
                .gradient_fill(GradientFill::two_color(135.0, purple, coral))
                .build(),
        );

        s.add_text(
            "LET'S WORK TOGETHER",
            TextOptionsBuilder::new()
                .bounds(1.0, 0.4, 8.0, 0.7)
                .font_size(36.0)
                .bold()
                .color(white)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_text(
            "I'm always open to new opportunities and collaborations",
            TextOptionsBuilder::new()
                .bounds(1.0, 1.2, 8.0, 0.5)
                .font_size(14.0)
                .color(white)
                .align(AlignH::Center)
                .valign(AlignV::Top)
                .build(),
        );

        // Contact items with shape icons
        let contact_area = deckmint::layout::CellRect {
            x: 1.0,
            y: 2.8,
            w: 8.0,
            h: 2.2,
        };
        let contact_rows = split_v(&contact_area, 4, 0.1);

        let contacts = [
            (ShapeType::Ellipse, "alex@morgandesign.com", "Email", coral),
            (ShapeType::RoundRect, "alexmorgan.design", "Portfolio", purple),
            (ShapeType::Diamond, "github.com/alexmorgan", "GitHub", teal),
            (ShapeType::Hexagon, "linkedin.com/in/alexmorgan", "LinkedIn", "#0077B5"),
        ];

        for (i, (shape, info, label, color)) in contacts.iter().enumerate() {
            let row = contact_rows[i];

            // Icon shape
            let icon_size = 0.38;
            let icon_x = row.x + 0.5;
            let icon_y = row.y + (row.h - icon_size) / 2.0;
            s.add_shape(
                shape.clone(),
                ShapeOptionsBuilder::new()
                    .bounds(icon_x, icon_y, icon_size, icon_size)
                    .fill_color(*color)
                    .build(),
            );

            // Label icon text
            let initial = label.chars().next().unwrap().to_string();
            s.add_text(
                &initial,
                TextOptionsBuilder::new()
                    .bounds(icon_x, icon_y, icon_size, icon_size)
                    .font_size(14.0)
                    .bold()
                    .color(white)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Label
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(row.x + 1.2, row.y, 1.5, row.h)
                    .font_size(11.0)
                    .bold()
                    .color(*color)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Contact info
            s.add_text(
                *info,
                TextOptionsBuilder::new()
                    .bounds(row.x + 2.8, row.y, 4.5, row.h)
                    .font_size(13.0)
                    .color(dark)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Bottom CTA
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(3.0, 5.0, 4.0, 0.4)
                .gradient_fill(GradientFill::two_color(0.0, coral, purple))
                .rect_radius(0.2)
                .build(),
        );
        s.add_text(
            "Get in Touch",
            TextOptionsBuilder::new()
                .bounds(3.0, 5.0, 4.0, 0.4)
                .font_size(14.0)
                .bold()
                .color(white)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    pres.write_to_file("34_creative_portfolio.pptx").unwrap();
    println!("Wrote 34_creative_portfolio.pptx");
}
