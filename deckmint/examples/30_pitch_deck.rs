//! 7-slide startup pitch deck for "NovaTech" — navy/blue palette with
//! SlideMasterDef branding, charts, tables, and grid layouts.

use deckmint::layout::{GridLayoutBuilder, split_v};
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::table::{TableCell, TableOptionsBuilder};
use deckmint::objects::text::{TextOptionsBuilder, TextRunBuilder};
use deckmint::{
    AlignH, AlignV, ChartOptionsBuilder, ChartSeries, ChartType, Presentation, ShapeType,
    SlideMasterDef,
};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "NovaTech Pitch Deck".to_string();

    // ── Color palette ─────────────────────────────────────────
    let navy = "#0A1628";
    let dark_blue = "#112240";
    let accent_blue = "#4472C4";
    let bright_blue = "#64B5F6";
    let light_blue = "#90CAF9";
    let white = "#FFFFFF";
    let muted = "#8FAADC";
    let green = "#70AD47";

    // ── Master slide: header bar + footer branding ────────────
    {
        let mut master = SlideMasterDef::new("NovaTech");
        master.set_background_color(navy);

        // Top accent bar
        master.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.0, 10.0, 0.05)
                .fill_color(accent_blue)
                .build(),
        );

        // Footer bar
        master.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 5.35, 10.0, 0.275)
                .fill_color(dark_blue)
                .build(),
        );

        // Footer text
        master.add_text(
            "NovaTech  |  Confidential",
            TextOptionsBuilder::new()
                .bounds(0.3, 5.35, 9.4, 0.275)
                .font_size(8.0)
                .color(muted)
                .valign(AlignV::Middle)
                .build(),
        );

        pres.define_master(master);
    }

    // ══════════════════════════════════════════════════════════
    // Slide 1: Title — company name + tagline
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        // Large background accent shape
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 1.8, 10.0, 2.0)
                .fill_color(dark_blue)
                .build(),
        );

        // Decorative line
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(3.0, 1.8, 4.0, 0.04)
                .fill_color(accent_blue)
                .build(),
        );

        s.add_text(
            "NOVATECH",
            TextOptionsBuilder::new()
                .bounds(1.0, 2.0, 8.0, 1.0)
                .font_size(52.0)
                .bold()
                .color(white)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_text(
            "Intelligent Automation for the Enterprise",
            TextOptionsBuilder::new()
                .bounds(1.0, 3.0, 8.0, 0.6)
                .font_size(20.0)
                .color(light_blue)
                .align(AlignH::Center)
                .valign(AlignV::Top)
                .build(),
        );

        // Decorative line
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(3.0, 3.7, 4.0, 0.04)
                .fill_color(accent_blue)
                .build(),
        );

        s.add_text(
            "Series A Fundraise  |  Q1 2026",
            TextOptionsBuilder::new()
                .bounds(1.0, 4.2, 8.0, 0.5)
                .font_size(14.0)
                .color(muted)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Problem — 3 pain points with icon shapes
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "THE PROBLEM",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(accent_blue)
                .build(),
        );

        s.add_text(
            "Enterprises are drowning in manual processes",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.5, 9.0, 0.6)
                .font_size(26.0)
                .bold()
                .color(white)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(3, 1, 0.3)
            .origin(0.5, 1.5)
            .container(9.0, 3.5)
            .build();

        let pain_points = [
            (ShapeType::Hexagon, "Slow Workflows", "87% of enterprises rely on\nmanual data entry, wasting\n20+ hours per week.", "#E74C3C"),
            (ShapeType::Diamond, "Fragmented Tools", "Teams juggle 12+ apps on\naverage, leading to silos\nand missed handoffs.", "#FFC000"),
            (ShapeType::Octagon, "Rising Costs", "Operational overhead grows\n15% year-over-year while\nmargins shrink.", "#ED7D31"),
        ];

        for (i, (shape, title, desc, color)) in pain_points.iter().enumerate() {
            let cell = grid.cell(i, 0);

            // Card background
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color(dark_blue)
                    .rect_radius(0.08)
                    .build(),
            );

            // Icon shape
            let icon_size = 0.6;
            let icon_x = cell.x + (cell.w - icon_size) / 2.0;
            s.add_shape(
                shape.clone(),
                ShapeOptionsBuilder::new()
                    .bounds(icon_x, cell.y + 0.3, icon_size, icon_size)
                    .fill_color(*color)
                    .build(),
            );

            // Title
            s.add_text(
                *title,
                TextOptionsBuilder::new()
                    .bounds(cell.x + 0.2, cell.y + 1.1, cell.w - 0.4, 0.5)
                    .font_size(16.0)
                    .bold()
                    .color(white)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Description
            s.add_text(
                *desc,
                TextOptionsBuilder::new()
                    .bounds(cell.x + 0.2, cell.y + 1.6, cell.w - 0.4, 1.4)
                    .font_size(11.0)
                    .color(muted)
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Solution — feature cards in grid
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "OUR SOLUTION",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(accent_blue)
                .build(),
        );

        s.add_text(
            "One platform to automate, connect, and scale",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.5, 9.0, 0.6)
                .font_size(26.0)
                .bold()
                .color(white)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(3, 2, 0.2)
            .origin(0.5, 1.4)
            .container(9.0, 3.8)
            .build();

        let features = [
            ("AI Workflows", "Auto-route tasks with\nML-powered decisions", accent_blue),
            ("Unified API", "Connect any system\nwith 200+ integrations", bright_blue),
            ("Real-time Sync", "Sub-second data sync\nacross all platforms", "#00B4D8"),
            ("Smart Analytics", "Actionable dashboards\nwith predictive insights", green),
            ("Enterprise SSO", "Bank-grade security\nwith SOC 2 compliance", "#9B59B6"),
            ("No-Code Builder", "Drag-and-drop workflow\ndesigner for any team", "#ED7D31"),
        ];

        for (idx, (title, desc, color)) in features.iter().enumerate() {
            let col = idx % 3;
            let row = idx / 3;
            let cell = grid.cell(col, row);

            // Card
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color(dark_blue)
                    .line_color(*color)
                    .line_width(1.5)
                    .rect_radius(0.06)
                    .build(),
            );

            // Accent bar
            s.add_shape(
                ShapeType::Rect,
                ShapeOptionsBuilder::new()
                    .bounds(cell.x, cell.y, cell.w, 0.04)
                    .fill_color(*color)
                    .build(),
            );

            let inner = cell.inset(0.15);
            s.add_text(
                *title,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y + 0.1, inner.w, 0.4)
                    .font_size(14.0)
                    .bold()
                    .color(white)
                    .valign(AlignV::Bottom)
                    .build(),
            );

            s.add_text(
                *desc,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y + 0.55, inner.w, 0.8)
                    .font_size(10.0)
                    .color(muted)
                    .valign(AlignV::Top)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Market — bar chart showing market growth
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "MARKET OPPORTUNITY",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(accent_blue)
                .build(),
        );

        s.add_text(
            "$48B enterprise automation market by 2028",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.5, 9.0, 0.6)
                .font_size(26.0)
                .bold()
                .color(white)
                .build(),
        );

        let years = vec!["2023", "2024", "2025", "2026", "2027", "2028"];
        let series = vec![
            ChartSeries::new("TAM ($B)", years, vec![22.0, 27.0, 32.0, 38.0, 43.0, 48.0]),
        ];

        s.add_chart(
            ChartType::Bar,
            series,
            ChartOptionsBuilder::new()
                .bounds(0.5, 1.3, 9.0, 3.8)
                .title("Total Addressable Market ($B)")
                .show_value()
                .chart_colors(vec![accent_blue])
                .val_axis_title("Revenue ($B)")
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 5: Business Model — pricing table
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "BUSINESS MODEL",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(accent_blue)
                .build(),
        );

        s.add_text(
            "SaaS pricing with land-and-expand strategy",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.5, 9.0, 0.6)
                .font_size(26.0)
                .bold()
                .color(white)
                .build(),
        );

        let hdr_bg = "#1B3A5C";
        let row_bg = "#0D2137";
        let header = vec![
            TableCell::new("Plan").bold().fill(hdr_bg).color(white).align(AlignH::Center),
            TableCell::new("Price / mo").bold().fill(hdr_bg).color(white).align(AlignH::Center),
            TableCell::new("Users").bold().fill(hdr_bg).color(white).align(AlignH::Center),
            TableCell::new("Features").bold().fill(hdr_bg).color(white).align(AlignH::Center),
        ];

        let rows = vec![
            header,
            vec![
                TableCell::new("Starter").bold().fill(row_bg).color(bright_blue).align(AlignH::Center),
                TableCell::new("$49").fill(row_bg).color(white).align(AlignH::Center),
                TableCell::new("Up to 10").fill(row_bg).color(muted).align(AlignH::Center),
                TableCell::new("Core workflows, 50 integrations").fill(row_bg).color(muted),
            ],
            vec![
                TableCell::new("Professional").bold().fill(row_bg).color(bright_blue).align(AlignH::Center),
                TableCell::new("$199").fill(row_bg).color(white).align(AlignH::Center),
                TableCell::new("Up to 50").fill(row_bg).color(muted).align(AlignH::Center),
                TableCell::new("AI workflows, analytics, SSO").fill(row_bg).color(muted),
            ],
            vec![
                TableCell::new("Enterprise").bold().fill(row_bg).color(bright_blue).align(AlignH::Center),
                TableCell::new("Custom").fill(row_bg).color(white).align(AlignH::Center),
                TableCell::new("Unlimited").fill(row_bg).color(muted).align(AlignH::Center),
                TableCell::new("Full platform, SLA, dedicated support").fill(row_bg).color(muted),
            ],
        ];

        s.add_table(
            rows,
            TableOptionsBuilder::new()
                .bounds(0.5, 1.4, 9.0, 2.5)
                .col_w(vec![2.0, 1.8, 1.8, 3.4])
                .font_size(12.0)
                .build(),
        );

        // Key metrics row below table
        let metrics_grid = GridLayoutBuilder::grid_n_m(3, 1, 0.3)
            .origin(0.5, 4.2)
            .container(9.0, 1.0)
            .build();

        let metrics = [
            ("120%", "Net Revenue Retention", green),
            ("$42K", "Average Contract Value", accent_blue),
            ("<6 mo", "Payback Period", bright_blue),
        ];

        for (i, (value, label, color)) in metrics.iter().enumerate() {
            let cell = metrics_grid.cell(i, 0);
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color(dark_blue)
                    .rect_radius(0.06)
                    .build(),
            );
            s.add_text_runs(
                vec![
                    TextRunBuilder::new(*value)
                        .font_size(22.0)
                        .bold()
                        .color(*color)
                        .build(),
                    TextRunBuilder::new(&format!("\n{}", label))
                        .font_size(10.0)
                        .color(muted)
                        .build(),
                ],
                TextOptionsBuilder::new()
                    .rect(cell.inset(0.1))
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 6: Team — 2x3 grid of team member cards
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "THE TEAM",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(accent_blue)
                .build(),
        );

        s.add_text(
            "Built by industry veterans",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.5, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color(white)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(3, 2, 0.2)
            .origin(0.5, 1.2)
            .container(9.0, 3.8)
            .build();

        let team = [
            ("Sarah Chen", "CEO & Co-founder", "Ex-Google, Stanford MBA"),
            ("James Park", "CTO & Co-founder", "Ex-AWS, 15yr distributed systems"),
            ("Maria Lopez", "VP Engineering", "Ex-Stripe, built payments infra"),
            ("David Kim", "VP Product", "Ex-Salesforce, 3 exits"),
            ("Priya Sharma", "VP Sales", "Ex-Datadog, scaled $0-50M ARR"),
            ("Alex Rivera", "VP Marketing", "Ex-HubSpot, growth expert"),
        ];

        for (idx, (name, role, bio)) in team.iter().enumerate() {
            let col = idx % 3;
            let row = idx / 3;
            let cell = grid.cell(col, row);

            // Card background
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color(dark_blue)
                    .rect_radius(0.06)
                    .build(),
            );

            // Avatar circle
            let avatar_size = 0.55;
            let avatar_x = cell.x + (cell.w - avatar_size) / 2.0;
            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(avatar_x, cell.y + 0.15, avatar_size, avatar_size)
                    .fill_color(accent_blue)
                    .build(),
            );

            // Initials in avatar
            let initials: String = name
                .split_whitespace()
                .map(|w| w.chars().next().unwrap())
                .collect();
            s.add_text(
                &initials,
                TextOptionsBuilder::new()
                    .bounds(avatar_x, cell.y + 0.15, avatar_size, avatar_size)
                    .font_size(16.0)
                    .bold()
                    .color(white)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Name
            s.add_text(
                *name,
                TextOptionsBuilder::new()
                    .bounds(cell.x + 0.1, cell.y + 0.8, cell.w - 0.2, 0.35)
                    .font_size(13.0)
                    .bold()
                    .color(white)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Role
            s.add_text(
                *role,
                TextOptionsBuilder::new()
                    .bounds(cell.x + 0.1, cell.y + 1.1, cell.w - 0.2, 0.3)
                    .font_size(10.0)
                    .bold()
                    .color(accent_blue)
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );

            // Bio
            s.add_text(
                *bio,
                TextOptionsBuilder::new()
                    .bounds(cell.x + 0.1, cell.y + 1.35, cell.w - 0.2, 0.4)
                    .font_size(9.0)
                    .color(muted)
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 7: CTA — contact info, large text
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        // Large background accent
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 1.0, 10.0, 3.5)
                .fill_color(dark_blue)
                .build(),
        );

        s.add_text(
            "LET'S BUILD THE FUTURE TOGETHER",
            TextOptionsBuilder::new()
                .bounds(1.0, 1.2, 8.0, 1.0)
                .font_size(36.0)
                .bold()
                .color(white)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(3.5, 2.3, 3.0, 0.04)
                .fill_color(accent_blue)
                .build(),
        );

        let contact_area = split_v(
            &deckmint::layout::CellRect { x: 1.5, y: 2.6, w: 7.0, h: 1.6 },
            3,
            0.1,
        );

        let contacts = [
            "sarah@novatech.io",
            "novatech.io/invest",
            "+1 (415) 555-0142",
        ];

        for (i, info) in contacts.iter().enumerate() {
            s.add_text(
                *info,
                TextOptionsBuilder::new()
                    .rect(contact_area[i])
                    .font_size(18.0)
                    .color(light_blue)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        s.add_text(
            "Raising $15M Series A  |  $8M committed",
            TextOptionsBuilder::new()
                .bounds(1.0, 4.6, 8.0, 0.5)
                .font_size(14.0)
                .bold()
                .color(green)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    pres.write_to_file("30_pitch_deck.pptx").unwrap();
    println!("Wrote 30_pitch_deck.pptx");
}
