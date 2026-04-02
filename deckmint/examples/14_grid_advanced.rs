//! Advanced grid layout: `.rect()`, `within()`, `sub_grid()`, spans, and nesting.

use deckmint::layout::{self, GridLayoutBuilder, GridTrack};
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();

    // ══════════════════════════════════════════════════════════
    // Slide 1: Dashboard layout with nested grids
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#F5F6FA");

        // Outer grid: sidebar + main area
        let outer = GridLayoutBuilder::sidebar_left(2.5, 0.15)
            .padding(0.2)
            .build();

        let sidebar = outer.cell(0, 0);
        let main_area = outer.cell(1, 0);

        // Sidebar background
        slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
            .rect(sidebar)
            .fill_color("#2C3E50")
            .rect_radius(0.08)
            .build());

        // Sidebar menu items
        let menu_items = ["Dashboard", "Analytics", "Reports", "Settings"];
        let menu_grid = GridLayoutBuilder::within(sidebar)
            .rows(vec![
                GridTrack::Inches(0.8),  // logo area
                GridTrack::Fr(1.0),      // menu items
            ])
            .padding(0.15)
            .build();

        // Logo area
        slide.add_text("DASH", TextOptionsBuilder::new()
            .rect(menu_grid.cell(0, 0))
            .font_size(22.0).bold()
            .color("#3498DB")
            .align(AlignH::Center).valign(AlignV::Middle)
            .build());

        // Menu items in the remaining space
        let items_area = menu_grid.cell(0, 1);
        let rows = layout::split_v(&items_area, menu_items.len(), 0.08);
        for (i, item) in menu_items.iter().enumerate() {
            let is_active = i == 0;
            let bg = if is_active { "#3498DB" } else { "#34495E" };
            slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
                .rect(rows[i])
                .fill_color(bg)
                .rect_radius(0.04)
                .build());
            slide.add_text(*item, TextOptionsBuilder::new()
                .rect(rows[i].inset_xy(0.15, 0.0))
                .font_size(12.0)
                .color("#FFFFFF")
                .valign(AlignV::Middle)
                .build());
        }

        // Main content: header row + 2x2 card grid
        let main_grid = GridLayoutBuilder::within(main_area)
            .rows(vec![GridTrack::Inches(0.7), GridTrack::Fr(1.0)])
            .row_gap(0.12)
            .build();

        slide.add_text("Dashboard Overview", TextOptionsBuilder::new()
            .rect(main_grid.cell(0, 0))
            .font_size(22.0).bold()
            .color("#2C3E50")
            .valign(AlignV::Bottom)
            .build());

        // 2x2 card grid nested inside the content cell
        let cards = main_grid.sub_grid(0, 1)
            .cols(vec![GridTrack::Fr(1.0); 2])
            .rows(vec![GridTrack::Fr(1.0); 2])
            .gap(0.12)
            .build();

        let card_data = [
            ("Users",    "12,847",  "+12%", "#3498DB"),
            ("Revenue",  "$84.2K",  "+8%",  "#2ECC71"),
            ("Orders",   "1,423",   "+23%", "#E67E22"),
            ("Uptime",   "99.97%",  "+0.1%","#9B59B6"),
        ];
        for (i, (label, value, delta, color)) in card_data.iter().enumerate() {
            let cell = cards.cell(i % 2, i / 2);

            // Card background
            slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
                .rect(cell)
                .fill_color("#FFFFFF")
                .rect_radius(0.06)
                .build());

            // Color accent bar at top of card
            let accent = deckmint::layout::CellRect {
                x: cell.x, y: cell.y, w: cell.w, h: 0.05,
            };
            slide.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
                .rect(accent)
                .fill_color(*color)
                .build());

            // Card content
            let inner = cell.inset(0.15);
            let (top, bot) = inner.halves_v(0.0);
            slide.add_text(*value, TextOptionsBuilder::new()
                .rect(top)
                .font_size(28.0).bold()
                .color("#2C3E50")
                .valign(AlignV::Bottom)
                .build());
            slide.add_text(&format!("{} {}", label, delta), TextOptionsBuilder::new()
                .rect(bot)
                .font_size(11.0)
                .color("#95A5A6")
                .valign(AlignV::Top)
                .build());
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Spanning cells — header + mixed layout
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FFFFFF");

        let grid = GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(2.0), GridTrack::Fr(1.0), GridTrack::Fr(1.0)])
            .rows(vec![
                GridTrack::Inches(0.8),  // title row
                GridTrack::Fr(1.0),      // main row
                GridTrack::Inches(1.2),  // bottom row
            ])
            .gap(0.12)
            .padding(0.3)
            .build();

        // Title spans all 3 columns
        let title_bar = grid.span(0, 0, 3, 1);
        slide.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
            .rect(title_bar)
            .fill_color("#1A1A2E")
            .build());
        slide.add_text("Project Status Report", TextOptionsBuilder::new()
            .rect(title_bar.inset_xy(0.2, 0.0))
            .font_size(24.0).bold()
            .color("#FFFFFF")
            .valign(AlignV::Middle)
            .build());

        // Left: large feature panel (spans 2 rows)
        let feature = grid.span(0, 1, 1, 2);
        slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
            .rect(feature)
            .fill_color("#16213E")
            .rect_radius(0.08)
            .build());
        let (ft, fb) = feature.inset(0.25).halves_v(0.1);
        slide.add_text("Sprint 24", TextOptionsBuilder::new()
            .rect(ft)
            .font_size(30.0).bold()
            .color("#E94560")
            .valign(AlignV::Bottom)
            .build());
        slide.add_text("14 stories completed\n3 in review\n1 blocked", TextOptionsBuilder::new()
            .rect(fb)
            .font_size(14.0)
            .color("#A0A0C0")
            .valign(AlignV::Top)
            .build());

        // Top-right: two small cards
        let card1 = grid.cell(1, 1);
        let card2 = grid.cell(2, 1);

        for (cell, label, value, color) in [
            (card1, "Velocity", "42 pts", "#0F3460"),
            (card2, "Burndown", "78%", "#533483"),
        ] {
            slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
                .rect(cell)
                .fill_color(color)
                .rect_radius(0.06)
                .build());
            let (t, b) = cell.inset(0.12).halves_v(0.0);
            slide.add_text(label, TextOptionsBuilder::new()
                .rect(t)
                .font_size(11.0)
                .color("#8899AA")
                .valign(AlignV::Bottom).align(AlignH::Center)
                .build());
            slide.add_text(value, TextOptionsBuilder::new()
                .rect(b)
                .font_size(24.0).bold()
                .color("#FFFFFF")
                .valign(AlignV::Top).align(AlignH::Center)
                .build());
        }

        // Bottom-right: spans 2 columns
        let timeline = grid.span(1, 2, 2, 1);
        slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
            .rect(timeline)
            .fill_color("#1A1A2E")
            .rect_radius(0.06)
            .build());
        slide.add_text("Timeline: On Track for July Release", TextOptionsBuilder::new()
            .rect(timeline)
            .font_size(14.0).bold()
            .color("#5BC0BE")
            .align(AlignH::Center).valign(AlignV::Middle)
            .build());
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Responsive-style layout with MinMax tracks
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#0B0C10");

        // MinMax ensures columns are at least 2" even if container shrinks
        let grid = GridLayoutBuilder::new()
            .cols(vec![
                GridTrack::MinMax { min: 2.0, max_fr: 1.0 },
                GridTrack::MinMax { min: 2.0, max_fr: 1.0 },
                GridTrack::MinMax { min: 2.0, max_fr: 1.0 },
            ])
            .rows(vec![GridTrack::Inches(0.7), GridTrack::Fr(1.0)])
            .gap(0.15)
            .padding(0.4)
            .build();

        // Header spans all columns
        let header = grid.span(0, 0, 3, 1);
        slide.add_text("Feature Comparison", TextOptionsBuilder::new()
            .rect(header)
            .font_size(24.0).bold()
            .color("#66FCF1")
            .valign(AlignV::Bottom)
            .build());

        let tiers = [
            ("Starter", "$9/mo", ["5 users", "10GB storage", "Email support"], "#1F2833"),
            ("Pro",     "$29/mo", ["25 users", "100GB storage", "Priority support"], "#45A29E"),
            ("Enterprise", "$99/mo", ["Unlimited", "1TB storage", "24/7 phone"], "#66FCF1"),
        ];

        for (col, (name, price, features, accent)) in tiers.iter().enumerate() {
            let cell = grid.cell(col, 1);

            // Card background
            let is_highlight = col == 1;
            let bg = if is_highlight { "#1F2833" } else { "#0B0C10" };
            slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
                .rect(cell)
                .fill_color(bg)
                .line_color(*accent).line_width(if is_highlight { 2.0 } else { 1.0 })
                .rect_radius(0.08)
                .build());

            // Split card into sections
            let inner = cell.inset(0.2);
            let sections = layout::split_v(&inner, 5, 0.08);

            // Tier name
            slide.add_text(*name, TextOptionsBuilder::new()
                .rect(sections[0])
                .font_size(20.0).bold()
                .color(*accent)
                .align(AlignH::Center).valign(AlignV::Middle)
                .build());

            // Price
            slide.add_text(*price, TextOptionsBuilder::new()
                .rect(sections[1])
                .font_size(16.0)
                .color("#C5C6C7")
                .align(AlignH::Center).valign(AlignV::Middle)
                .build());

            // Feature list
            for (i, feat) in features.iter().enumerate() {
                slide.add_text(*feat, TextOptionsBuilder::new()
                    .rect(sections[2 + i])
                    .font_size(11.0)
                    .color("#C5C6C7")
                    .align(AlignH::Center).valign(AlignV::Middle)
                    .build());
            }
        }
    }

    pres.write_to_file("14_grid_advanced.pptx").unwrap();
    println!("Wrote 14_grid_advanced.pptx");
}
