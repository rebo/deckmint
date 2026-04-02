//! 3-slide data dashboard — dark theme with KPI cards, line chart, pie chart,
//! bar chart, and full-width area chart. Uses grid layout and master slide.

use deckmint::layout::GridLayoutBuilder;
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{
    AlignH, AlignV, ChartOptionsBuilder, ChartSeries, ChartType, LegendPos, Presentation,
    ShapeType, SlideMasterDef,
};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Data Dashboard".to_string();

    // ── Colors ────────────────────────────────────────────────
    let bg = "#1A1A2E";
    let card_bg = "#16213E";
    let accent = "#0F3460";
    let cyan = "#00D2FF";
    let green = "#00E676";
    let orange = "#FF9100";
    let pink = "#E94560";
    let white = "#FFFFFF";
    let muted = "#607080";

    // ── Master slide ──────────────────────────────────────────
    {
        let mut master = SlideMasterDef::new("Dashboard");
        master.set_background_color(bg);

        // Top accent bar
        master.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.0, 10.0, 0.04)
                .fill_color(cyan)
                .build(),
        );

        // Bottom bar
        master.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 5.4, 10.0, 0.225)
                .fill_color(accent)
                .build(),
        );
        master.add_text(
            "Analytics Dashboard  |  Real-time Metrics",
            TextOptionsBuilder::new()
                .bounds(0.3, 5.4, 9.4, 0.225)
                .font_size(8.0)
                .color(muted)
                .valign(AlignV::Middle)
                .build(),
        );

        pres.define_master(master);
    }

    // ══════════════════════════════════════════════════════════
    // Slide 1: KPI cards (4 across top) + line chart below
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        // Title
        s.add_text(
            "Performance Overview",
            TextOptionsBuilder::new()
                .bounds(0.4, 0.15, 9.2, 0.4)
                .font_size(20.0)
                .bold()
                .color(white)
                .build(),
        );

        // KPI cards row
        let kpi_grid = GridLayoutBuilder::grid_n_m(4, 1, 0.15)
            .origin(0.4, 0.6)
            .container(9.2, 1.4)
            .build();

        let kpis = [
            ("$2.4M", "Revenue", "+18%", cyan),
            ("12,847", "Active Users", "+24%", green),
            ("$186", "Avg. Order", "+7%", orange),
            ("3.2%", "Churn Rate", "-0.5%", pink),
        ];

        for (i, (value, label, delta, color)) in kpis.iter().enumerate() {
            let cell = kpi_grid.cell(i, 0);

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color(card_bg)
                    .rect_radius(0.06)
                    .build(),
            );

            // Left accent bar
            s.add_shape(
                ShapeType::Rect,
                ShapeOptionsBuilder::new()
                    .bounds(cell.x, cell.y, 0.04, cell.h)
                    .fill_color(*color)
                    .build(),
            );

            let inner = cell.inset_xy(0.15, 0.1);

            s.add_text(
                *value,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y, inner.w, 0.55)
                    .font_size(24.0)
                    .bold()
                    .color(*color)
                    .valign(AlignV::Bottom)
                    .build(),
            );

            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y + 0.55, inner.w * 0.6, 0.35)
                    .font_size(10.0)
                    .color(muted)
                    .valign(AlignV::Top)
                    .build(),
            );

            s.add_text(
                *delta,
                TextOptionsBuilder::new()
                    .bounds(inner.x + inner.w * 0.6, inner.y + 0.55, inner.w * 0.4, 0.35)
                    .font_size(10.0)
                    .bold()
                    .color(if delta.starts_with('-') { pink } else { green })
                    .align(AlignH::Right)
                    .valign(AlignV::Top)
                    .build(),
            );
        }

        // Line chart below KPIs
        let months = vec![
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ];
        let series = vec![
            ChartSeries::new(
                "Revenue ($K)",
                months.clone(),
                vec![145.0, 162.0, 158.0, 178.0, 195.0, 210.0, 198.0, 225.0, 240.0, 258.0, 272.0, 290.0],
            ),
            ChartSeries::new(
                "Expenses ($K)",
                months,
                vec![120.0, 128.0, 135.0, 130.0, 142.0, 148.0, 155.0, 160.0, 162.0, 170.0, 175.0, 182.0],
            ),
        ];

        s.add_chart(
            ChartType::Line,
            series,
            ChartOptionsBuilder::new()
                .bounds(0.4, 2.15, 9.2, 3.1)
                .title("Revenue vs Expenses (12-Month Trend)")
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(vec![cyan, orange])
                .line_smooth()
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Pie chart + bar chart side by side
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "Segment Analysis",
            TextOptionsBuilder::new()
                .bounds(0.4, 0.15, 9.2, 0.4)
                .font_size(20.0)
                .bold()
                .color(white)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(2, 1, 0.2)
            .origin(0.4, 0.65)
            .container(9.2, 4.6)
            .build();

        // Left: Pie chart — traffic sources
        let pie_series = vec![ChartSeries::new(
            "Traffic",
            vec!["Organic", "Paid", "Social", "Referral", "Direct"],
            vec![38.0, 24.0, 18.0, 12.0, 8.0],
        )];

        let left = grid.cell(0, 0);

        // Card background for pie
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .rect(left)
                .fill_color(card_bg)
                .rect_radius(0.08)
                .build(),
        );

        s.add_chart(
            ChartType::Pie,
            pie_series,
            ChartOptionsBuilder::new()
                .rect(left.inset(0.1))
                .title("Traffic Sources")
                .show_value()
                .chart_colors(vec![cyan, orange, green, pink, "#9B59B6"])
                .build(),
        );

        // Right: Bar chart — conversion by channel
        let channels = vec!["Organic", "Paid", "Social", "Referral", "Direct"];
        let bar_series = vec![
            ChartSeries::new("Conversion %", channels, vec![4.2, 3.8, 2.5, 5.1, 6.3]),
        ];

        let right = grid.cell(1, 0);

        // Card background for bar
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .rect(right)
                .fill_color(card_bg)
                .rect_radius(0.08)
                .build(),
        );

        s.add_chart(
            ChartType::Bar,
            bar_series,
            ChartOptionsBuilder::new()
                .rect(right.inset(0.1))
                .title("Conversion by Channel (%)")
                .show_value()
                .chart_colors(vec![cyan])
                .val_axis_title("Conversion %")
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Full-width area chart
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "Growth Trajectory",
            TextOptionsBuilder::new()
                .bounds(0.4, 0.15, 9.2, 0.4)
                .font_size(20.0)
                .bold()
                .color(white)
                .build(),
        );

        // Summary cards at top
        let summary_grid = GridLayoutBuilder::grid_n_m(3, 1, 0.15)
            .origin(0.4, 0.6)
            .container(9.2, 0.7)
            .build();

        let summaries = [
            ("Total Users: 48.2K", cyan),
            ("Growth Rate: 12.4% MoM", green),
            ("Projected EOY: 92K", orange),
        ];

        for (i, (text, color)) in summaries.iter().enumerate() {
            let cell = summary_grid.cell(i, 0);
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color(card_bg)
                    .rect_radius(0.04)
                    .build(),
            );
            s.add_text(
                *text,
                TextOptionsBuilder::new()
                    .rect(cell)
                    .font_size(11.0)
                    .bold()
                    .color(*color)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Full-width area chart
        let months = vec![
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ];

        let series = vec![
            ChartSeries::new(
                "Web Users (K)",
                months.clone(),
                vec![12.0, 14.5, 16.2, 18.8, 21.0, 24.5, 27.0, 30.2, 33.8, 37.5, 42.0, 48.2],
            ),
            ChartSeries::new(
                "Mobile Users (K)",
                months.clone(),
                vec![5.0, 6.8, 8.2, 10.0, 12.5, 15.0, 17.8, 20.5, 23.0, 26.2, 30.0, 35.5],
            ),
            ChartSeries::new(
                "API Calls (K)",
                months,
                vec![2.0, 3.2, 4.5, 6.0, 8.0, 10.5, 13.0, 16.0, 19.5, 23.0, 27.0, 32.0],
            ),
        ];

        // Chart background card
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(0.4, 1.45, 9.2, 3.8)
                .fill_color(card_bg)
                .rect_radius(0.08)
                .build(),
        );

        s.add_chart(
            ChartType::Area,
            series,
            ChartOptionsBuilder::new()
                .bounds(0.5, 1.55, 9.0, 3.6)
                .title("User Growth by Platform (Thousands)")
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(vec![cyan, green, orange])
                .cat_axis_title("Month")
                .val_axis_title("Users (K)")
                .build(),
        );
    }

    pres.write_to_file("33_dashboard.pptx").unwrap();
    println!("Wrote 33_dashboard.pptx");
}
