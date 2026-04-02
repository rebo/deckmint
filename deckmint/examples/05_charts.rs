//! Bar, line, and pie charts using the chart builder API.

use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{
    AlignH, BarDir, ChartOptionsBuilder, ChartSeries, ChartType, LegendPos, Presentation,
};

fn main() {
    let mut pres = Presentation::new();

    let labels = vec!["Q1", "Q2", "Q3", "Q4"];
    let series = vec![
        ChartSeries::new("Product A", labels.clone(), vec![12.0, 19.0, 14.0, 22.0]),
        ChartSeries::new("Product B", labels.clone(), vec![8.0, 15.0, 21.0, 17.0]),
    ];

    // Slide 1: Column chart
    let s = pres.add_slide();
    s.add_text("Column Chart", TextOptionsBuilder::new()
        .pos(0.5, 0.2).size(9.0, 0.7).font_size(24.0).bold().build());
    s.add_chart(
        ChartType::Bar,
        series.clone(),
        ChartOptionsBuilder::new()
            .pos(0.5, 1.0).size(9.0, 4.3)
            .title("Quarterly Revenue")
            .show_value()
            .chart_colors(vec!["4472C4", "ED7D31"])
            .build(),
    );

    // Slide 2: Horizontal bar chart
    let s = pres.add_slide();
    s.add_text("Horizontal Bar", TextOptionsBuilder::new()
        .pos(0.5, 0.2).size(9.0, 0.7).font_size(24.0).bold().build());
    s.add_chart(
        ChartType::Bar,
        series.clone(),
        ChartOptionsBuilder::new()
            .pos(0.5, 1.0).size(9.0, 4.3)
            .title("Quarterly Revenue (Bar)")
            .bar_dir(BarDir::Bar)
            .legend_pos(LegendPos::Bottom)
            .build(),
    );

    // Slide 3: Pie chart
    let pie_series = vec![
        ChartSeries::new("Market", vec!["US", "EU", "Asia", "Other"], vec![42.0, 28.0, 22.0, 8.0]),
    ];
    let s = pres.add_slide();
    s.add_text("Pie Chart", TextOptionsBuilder::new()
        .pos(0.5, 0.2).size(9.0, 0.7).font_size(24.0).bold().build());
    s.add_chart(
        ChartType::Pie,
        pie_series,
        ChartOptionsBuilder::new()
            .pos(1.0, 1.0).size(8.0, 4.3)
            .title("Market Share")
            .show_value()
            .chart_colors(vec!["4472C4", "ED7D31", "70AD47", "FFC000"])
            .build(),
    );

    // Slide 4: Line chart
    let s = pres.add_slide();
    s.add_text("Line Chart", TextOptionsBuilder::new()
        .pos(0.5, 0.2).size(9.0, 0.7).font_size(24.0).bold().build());
    s.add_chart(
        ChartType::Line,
        series,
        ChartOptionsBuilder::new()
            .pos(0.5, 1.0).size(9.0, 4.3)
            .title("Trend Analysis")
            .line_smooth()
            .cat_axis_title("Quarter")
            .val_axis_title("Revenue ($M)")
            .build(),
    );

    pres.write_to_file("05_charts.pptx").unwrap();
    println!("Wrote 05_charts.pptx");
}
