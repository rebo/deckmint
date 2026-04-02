//! Master slide ergonomics: convenience methods, grid integration, and promotion.

use deckmint::layout::{GridLayoutBuilder, GridTrack};
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, Presentation, ShapeType, SlideMasterDef};

fn main() {
    // ══════════════════════════════════════════════════════════
    // Part 1: Build a master directly with convenience methods
    // ══════════════════════════════════════════════════════════

    let mut pres = Presentation::new();

    // Define a grid for the master chrome: header + content + footer
    let master_grid = GridLayoutBuilder::new()
        .cols(vec![GridTrack::Fr(1.0)])
        .rows(vec![
            GridTrack::Inches(0.6),  // header bar
            GridTrack::Fr(1.0),      // content (not used on master)
            GridTrack::Inches(0.35), // footer bar
        ])
        .build();

    let header = master_grid.cell(0, 0);
    let footer = master_grid.cell(0, 2);

    let mut master = SlideMasterDef::new("Corporate");
    master.set_background_color("#F2F2F2");

    // Dark header bar
    master.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
        .rect(header)
        .fill_color("#1B2A4A")
        .build());

    // Company name in header
    master.add_text("ACME Corporation", TextOptionsBuilder::new()
        .rect(header.inset_xy(0.3, 0.0))
        .font_size(16.0).bold()
        .color("#FFFFFF")
        .valign(AlignV::Middle)
        .build());

    // Subtle footer bar
    master.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
        .rect(footer)
        .fill_color("#E0E0E0")
        .build());

    // Footer text
    master.add_text("Confidential", TextOptionsBuilder::new()
        .rect(footer.inset_xy(0.3, 0.0))
        .font_size(9.0).italic()
        .color("#888888")
        .valign(AlignV::Middle)
        .align(AlignH::Right)
        .build());

    pres.define_master(master);

    // ── Slide 1: Content uses the same grid for consistent positioning ──
    {
        let content = master_grid.cell(0, 1);
        let inner = GridLayoutBuilder::within(content)
            .cols(vec![GridTrack::Fr(1.0)])
            .rows(vec![GridTrack::Inches(0.8), GridTrack::Fr(1.0)])
            .row_gap(0.15)
            .padding(0.4)
            .build();

        let slide = pres.add_slide();
        slide.add_text("Q4 Results Overview", TextOptionsBuilder::new()
            .rect(inner.cell(0, 0))
            .font_size(28.0).bold()
            .color("#1B2A4A")
            .valign(AlignV::Bottom)
            .build());

        // 3-column stats in the content area
        let stats_grid = inner.sub_grid(0, 1)
            .cols(vec![GridTrack::Fr(1.0); 3])
            .gap(0.2)
            .build();

        let stats = [
            ("Revenue", "$4.2M", "#4472C4"),
            ("Growth",  "+18%",  "#70AD47"),
            ("Margin",  "32%",   "#ED7D31"),
        ];
        for (col, (label, value, color)) in stats.iter().enumerate() {
            let cell = stats_grid.cell(col, 0);

            slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
                .rect(cell)
                .fill_color(*color)
                .rect_radius(0.08)
                .build());

            let (top, bottom) = cell.inset(0.15).halves_v(0.1);
            slide.add_text(*value, TextOptionsBuilder::new()
                .rect(top)
                .font_size(36.0).bold()
                .color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Bottom)
                .build());
            slide.add_text(*label, TextOptionsBuilder::new()
                .rect(bottom)
                .font_size(14.0)
                .color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Top)
                .build());
        }
    }

    // ── Slide 2: Two-column layout in the content area ──
    {
        let content = master_grid.cell(0, 1);
        let inner = GridLayoutBuilder::within(content)
            .cols(vec![GridTrack::Fr(1.0), GridTrack::Fr(1.0)])
            .rows(vec![GridTrack::Inches(0.7), GridTrack::Fr(1.0)])
            .col_gap(0.3)
            .row_gap(0.15)
            .padding(0.4)
            .build();

        let slide = pres.add_slide();

        // Title spans both columns
        let title = inner.span(0, 0, 2, 1);
        slide.add_text("Strategic Priorities", TextOptionsBuilder::new()
            .rect(title)
            .font_size(28.0).bold()
            .color("#1B2A4A")
            .valign(AlignV::Bottom)
            .build());

        // Left column: bullet items
        let left = inner.cell(0, 1);
        let items = ["Expand market share", "Launch new product line", "Improve retention"];
        let rows = deckmint::layout::split_v(&left, items.len(), 0.1);
        for (i, item) in items.iter().enumerate() {
            slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
                .rect(rows[i])
                .fill_color("#FFFFFF")
                .line_color("#D0D0D0").line_width(1.0)
                .rect_radius(0.06)
                .build());
            slide.add_text(&format!("{}. {}", i + 1, item), TextOptionsBuilder::new()
                .rect(rows[i].inset_xy(0.2, 0.0))
                .font_size(14.0)
                .color("#333333")
                .valign(AlignV::Middle)
                .build());
        }

        // Right column: single highlight box
        let right = inner.cell(1, 1);
        slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
            .rect(right)
            .fill_color("#1B2A4A")
            .rect_radius(0.08)
            .build());
        let (top, bot) = right.inset(0.3).halves_v(0.2);
        slide.add_text("Target", TextOptionsBuilder::new()
            .rect(top)
            .font_size(16.0)
            .color("#8FAADC")
            .align(AlignH::Center).valign(AlignV::Bottom)
            .build());
        slide.add_text("$10M ARR", TextOptionsBuilder::new()
            .rect(bot)
            .font_size(40.0).bold()
            .color("#FFFFFF")
            .align(AlignH::Center).valign(AlignV::Top)
            .build());
    }

    // ══════════════════════════════════════════════════════════
    // Part 2: Promote a slide to master
    // ══════════════════════════════════════════════════════════

    let mut pres2 = Presentation::new();

    // Build the master visually as a regular slide
    {
        let s = pres2.add_slide();
        s.set_background_color("#0D1B2A");

        // Gradient-like effect with overlapping shapes
        s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
            .bounds(0.0, 0.0, 10.0, 0.08)
            .fill_color("#4CC9F0")
            .build());
        s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
            .bounds(0.0, 5.545, 10.0, 0.08)
            .fill_color("#4CC9F0")
            .build());

        // Side accent
        s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
            .bounds(0.0, 0.08, 0.04, 5.465)
            .fill_color("#7209B7")
            .build());

        s.add_text("TechCo", TextOptionsBuilder::new()
            .bounds(0.3, 5.1, 2.0, 0.4)
            .font_size(10.0).bold()
            .color("#4CC9F0")
            .build());
    }

    // Promote slide 0 to master, then add real content slides
    let dark_master = pres2.promote_slide_to_master(0, "Dark Theme");
    pres2.define_master(dark_master);

    {
        let slide = pres2.add_slide();
        slide.add_text("Welcome to TechCo", TextOptionsBuilder::new()
            .bounds(1.0, 1.5, 8.0, 1.5)
            .font_size(40.0).bold()
            .color("#FFFFFF")
            .align(AlignH::Center)
            .build());
        slide.add_text("Building the future, one slide at a time", TextOptionsBuilder::new()
            .bounds(1.0, 3.0, 8.0, 1.0)
            .font_size(18.0)
            .color("#A0A0A0")
            .align(AlignH::Center)
            .build());
    }

    {
        let slide = pres2.add_slide();
        let grid = GridLayoutBuilder::grid_n_m(2, 2, 0.25)
            .origin(0.8, 0.8).container(8.4, 4.0)
            .build();

        let items = [
            ("Cloud", "Scale infinitely", "#4CC9F0"),
            ("AI/ML", "Smart automation", "#7209B7"),
            ("Security", "Zero-trust model", "#F72585"),
            ("DevOps", "Ship faster", "#4361EE"),
        ];
        for (i, (title, desc, color)) in items.iter().enumerate() {
            let cell = grid.cell(i % 2, i / 2);
            slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
                .rect(cell)
                .fill_color("#1B2838")
                .line_color(*color).line_width(2.0)
                .rect_radius(0.1)
                .build());
            let (top, bot) = cell.inset(0.2).halves_v(0.05);
            slide.add_text(*title, TextOptionsBuilder::new()
                .rect(top)
                .font_size(22.0).bold()
                .color(*color)
                .valign(AlignV::Bottom)
                .build());
            slide.add_text(*desc, TextOptionsBuilder::new()
                .rect(bot)
                .font_size(13.0)
                .color("#A0A0A0")
                .valign(AlignV::Top)
                .build());
        }
    }

    // ── Write both presentations ────────────────────────
    pres.write_to_file("13_master_slides.pptx").unwrap();
    println!("Wrote 13_master_slides.pptx (corporate master with grid layout)");

    pres2.write_to_file("13_master_promoted.pptx").unwrap();
    println!("Wrote 13_master_promoted.pptx (promoted slide as dark theme master)");
}
