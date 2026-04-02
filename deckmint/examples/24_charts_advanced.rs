//! Advanced chart features: stacked bars, horizontal bars, doughnut, radar, area, and scatter.

use deckmint::layout::GridLayoutBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{
    AlignH, BarDir, BarGrouping, ChartOptionsBuilder, ChartSeries, ChartType, LegendPos,
    Presentation,
};

fn main() {
    let mut pres = Presentation::new();

    // Shared data
    let categories = vec!["Q1", "Q2", "Q3", "Q4"];
    let series_revenue = vec![
        ChartSeries::new("North", categories.clone(), vec![45.0, 52.0, 48.0, 61.0]),
        ChartSeries::new("South", categories.clone(), vec![32.0, 38.0, 41.0, 35.0]),
        ChartSeries::new("East", categories.clone(), vec![28.0, 31.0, 36.0, 42.0]),
        ChartSeries::new("West", categories.clone(), vec![19.0, 24.0, 22.0, 29.0]),
    ];

    let chart_colors = vec!["#4472C4", "#ED7D31", "#70AD47", "#FFC000"];

    // ══════════════════════════════════════════════════════════
    // Slide 1: Stacked bar chart and percent-stacked bar chart
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FFFFFF");

        slide.add_text(
            "Stacked & Percent-Stacked Charts",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(22.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(2, 1, 0.3)
            .origin(0.3, 0.7)
            .container(9.4, 4.7)
            .build();

        // Stacked bar chart (left)
        let left = grid.cell(0, 0);
        slide.add_chart(
            ChartType::Bar,
            series_revenue.clone(),
            ChartOptionsBuilder::new()
                .rect(left)
                .title("Revenue by Region (Stacked)")
                .bar_grouping(BarGrouping::Stacked)
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(chart_colors.clone())
                .val_axis_title("Revenue ($M)")
                .build(),
        );

        // Percent-stacked bar chart (right)
        let right = grid.cell(1, 0);
        slide.add_chart(
            ChartType::Bar,
            series_revenue.clone(),
            ChartOptionsBuilder::new()
                .rect(right)
                .title("Revenue Share (Percent Stacked)")
                .bar_grouping(BarGrouping::PercentStacked)
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(chart_colors.clone())
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Horizontal bar chart
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FFFFFF");

        slide.add_text(
            "Horizontal Bar Charts",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(22.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        let departments = vec!["Engineering", "Marketing", "Sales", "Design", "Operations", "Support"];
        let dept_series = vec![
            ChartSeries::new("Budget", departments.clone(), vec![120.0, 85.0, 95.0, 60.0, 75.0, 50.0]),
            ChartSeries::new("Actual", departments.clone(), vec![115.0, 92.0, 88.0, 55.0, 80.0, 48.0]),
        ];

        let grid = GridLayoutBuilder::grid_n_m(2, 1, 0.3)
            .origin(0.3, 0.7)
            .container(9.4, 4.7)
            .build();

        // Clustered horizontal bar (left)
        let left = grid.cell(0, 0);
        slide.add_chart(
            ChartType::Bar,
            dept_series.clone(),
            ChartOptionsBuilder::new()
                .rect(left)
                .title("Budget vs Actual ($K)")
                .bar_dir(BarDir::Bar)
                .show_value()
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(vec!["#5B9BD5", "#ED7D31"])
                .cat_axis_title("Department")
                .val_axis_title("Amount ($K)")
                .build(),
        );

        // Stacked horizontal bar (right)
        let right = grid.cell(1, 0);
        slide.add_chart(
            ChartType::Bar,
            dept_series,
            ChartOptionsBuilder::new()
                .rect(right)
                .title("Combined Spending")
                .bar_dir(BarDir::Bar)
                .bar_grouping(BarGrouping::Stacked)
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(vec!["#5B9BD5", "#ED7D31"])
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Doughnut chart and Radar chart
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FFFFFF");

        slide.add_text(
            "Doughnut & Radar Charts",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(22.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(2, 1, 0.3)
            .origin(0.3, 0.7)
            .container(9.4, 4.7)
            .build();

        // Doughnut chart (left)
        let doughnut_series = vec![
            ChartSeries::new(
                "Market Share",
                vec!["Cloud", "On-Prem", "Hybrid", "Edge", "Other"],
                vec![35.0, 25.0, 20.0, 12.0, 8.0],
            ),
        ];
        let left = grid.cell(0, 0);
        slide.add_chart(
            ChartType::Doughnut,
            doughnut_series,
            ChartOptionsBuilder::new()
                .rect(left)
                .title("Infrastructure Mix")
                .hole_size(55)
                .show_value()
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(vec!["#4472C4", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5"])
                .build(),
        );

        // Radar chart (right)
        let skills = vec!["Frontend", "Backend", "DevOps", "Security", "Testing", "Design"];
        let radar_series = vec![
            ChartSeries::new("Team A", skills.clone(), vec![9.0, 7.0, 6.0, 8.0, 7.0, 5.0]),
            ChartSeries::new("Team B", skills.clone(), vec![6.0, 9.0, 8.0, 5.0, 8.0, 4.0]),
            ChartSeries::new("Team C", skills.clone(), vec![7.0, 5.0, 7.0, 6.0, 9.0, 8.0]),
        ];
        let right = grid.cell(1, 0);
        slide.add_chart(
            ChartType::Radar,
            radar_series,
            ChartOptionsBuilder::new()
                .rect(right)
                .title("Team Skill Profiles")
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(vec!["#4472C4", "#ED7D31", "#70AD47"])
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Area chart and Scatter chart
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FFFFFF");

        slide.add_text(
            "Area & Scatter Charts",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(22.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(2, 1, 0.3)
            .origin(0.3, 0.7)
            .container(9.4, 4.7)
            .build();

        // Area chart (left)
        let months = vec!["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        let area_series = vec![
            ChartSeries::new("Website",  months.clone(), vec![12.0, 15.0, 18.0, 22.0, 28.0, 35.0, 42.0, 38.0, 45.0, 50.0, 55.0, 62.0]),
            ChartSeries::new("Mobile",   months.clone(), vec![5.0, 8.0, 12.0, 15.0, 20.0, 25.0, 30.0, 35.0, 40.0, 48.0, 52.0, 60.0]),
            ChartSeries::new("API",      months.clone(), vec![2.0, 3.0, 5.0, 7.0, 10.0, 14.0, 18.0, 22.0, 28.0, 32.0, 38.0, 45.0]),
        ];
        let left = grid.cell(0, 0);
        slide.add_chart(
            ChartType::Area,
            area_series,
            ChartOptionsBuilder::new()
                .rect(left)
                .title("Traffic Growth (K visits)")
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(vec!["#4472C4", "#ED7D31", "#70AD47"])
                .cat_axis_title("Month")
                .val_axis_title("Visitors (K)")
                .build(),
        );

        // Scatter chart (right)
        let x_vals: Vec<&str> = (1..=15).map(|i| match i {
            1 => "1", 2 => "2", 3 => "3", 4 => "4", 5 => "5",
            6 => "6", 7 => "7", 8 => "8", 9 => "9", 10 => "10",
            11 => "11", 12 => "12", 13 => "13", 14 => "14", 15 => "15",
            _ => "",
        }).collect();
        let scatter_series = vec![
            ChartSeries::new(
                "Study Hours vs Score",
                x_vals.clone(),
                vec![42.0, 55.0, 48.0, 62.0, 70.0, 65.0, 78.0, 72.0, 85.0, 80.0, 88.0, 92.0, 85.0, 95.0, 91.0],
            ),
            ChartSeries::new(
                "Practice Hours vs Score",
                x_vals,
                vec![38.0, 45.0, 52.0, 58.0, 55.0, 68.0, 72.0, 75.0, 70.0, 82.0, 78.0, 88.0, 90.0, 86.0, 95.0],
            ),
        ];
        let right = grid.cell(1, 0);
        slide.add_chart(
            ChartType::Scatter,
            scatter_series,
            ChartOptionsBuilder::new()
                .rect(right)
                .title("Hours vs Test Score")
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(vec!["#E74C3C", "#3498DB"])
                .cat_axis_title("Hours")
                .val_axis_title("Score")
                .no_grid_lines()
                .build(),
        );
    }

    pres.write_to_file("24_charts_advanced.pptx").unwrap();
    println!("Wrote 24_charts_advanced.pptx");
}
