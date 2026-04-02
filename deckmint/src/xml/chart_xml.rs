use crate::enums::ChartType;
use crate::objects::chart::{BarDir, BarGrouping, ChartObject, ChartSeries, DEFAULT_CHART_COLORS};
use crate::utils::encode_xml_entities;

/// Chart-level data labels block (all-off defaults). Required inside every
/// chart element after the series, before <c:axId> or <c:firstSliceAng>.
const CHART_LEVEL_DLBLS: &str = "\
<c:dLbls>\
<c:showLegendKey val=\"0\"/>\
<c:showVal val=\"0\"/>\
<c:showCatName val=\"0\"/>\
<c:showSerName val=\"0\"/>\
<c:showPercent val=\"0\"/>\
<c:showBubbleSize val=\"0\"/>\
</c:dLbls>";

// ─────────────────────────────────────────────────────────────
// Top-level entry: generate ppt/charts/chartN.xml content
// ─────────────────────────────────────────────────────────────

pub fn gen_xml_chart(chart: &ChartObject) -> String {
    let opts = &chart.options;
    let mut s = String::new();
    s.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
    s.push_str(
        "<c:chartSpace \
xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" \
xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
    );
    s.push_str("<c:lang val=\"en-US\"/>");
    s.push_str("<c:style val=\"2\"/>");
    s.push_str("<c:chart>");

    // Title
    if let Some(ref title) = opts.title {
        s.push_str(&gen_xml_chart_title(title));
    }
    s.push_str("<c:autoTitleDeleted val=\"0\"/>");

    // Plot area
    s.push_str("<c:plotArea><c:layout/>");
    s.push_str(&gen_xml_plot_area(chart));
    s.push_str("</c:plotArea>");

    // Legend
    if opts.show_legend {
        s.push_str(&format!(
            "<c:legend><c:legendPos val=\"{}\"/><c:overlay val=\"0\"/></c:legend>",
            opts.legend_pos.as_ooxml()
        ));
    }
    s.push_str("<c:plotVisOnly val=\"1\"/>");
    s.push_str("</c:chart>");

    // Chart area style: no fill, no border
    s.push_str("<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>");
    s.push_str("</c:chartSpace>");
    s
}

// ─────────────────────────────────────────────────────────────
// Slide-level graphicFrame element (inserted into <p:spTree>)
// ─────────────────────────────────────────────────────────────

pub fn gen_xml_chart_frame(chart: &ChartObject, obj_id: usize, pres: &crate::presentation::Presentation) -> String {
    use crate::utils::inch_to_emu;
    let layout = &pres.layout;
    let pos = &chart.options.position;
    let x = pos.x.as_ref().map(|c| c.to_emu(layout.width)).unwrap_or_else(|| inch_to_emu(1.0));
    let y = pos.y.as_ref().map(|c| c.to_emu(layout.height)).unwrap_or_else(|| inch_to_emu(1.5));
    let cx = pos.w.as_ref().map(|c| c.to_emu(layout.width)).unwrap_or_else(|| inch_to_emu(8.0));
    let cy = pos.h.as_ref().map(|c| c.to_emu(layout.height)).unwrap_or_else(|| inch_to_emu(5.0));
    let rid = chart.chart_rid;
    let name = encode_xml_entities(&chart.object_name);

    format!(
        "<p:graphicFrame>\
<p:nvGraphicFramePr>\
<p:cNvPr id=\"{obj_id}\" name=\"{name}\"/>\
<p:cNvGraphicFramePr/>\
<p:nvPr/>\
</p:nvGraphicFramePr>\
<p:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></p:xfrm>\
<a:graphic>\
<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\
<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId{rid}\"/>\
</a:graphicData>\
</a:graphic>\
</p:graphicFrame>"
    )
}

// ─────────────────────────────────────────────────────────────
// Chart title XML
// ─────────────────────────────────────────────────────────────

fn gen_xml_chart_title(title: &str) -> String {
    let t = encode_xml_entities(title);
    format!(
        "<c:title>\
<c:tx><c:rich>\
<a:bodyPr/><a:lstStyle/>\
<a:p><a:r><a:rPr lang=\"en-US\" dirty=\"0\"/><a:t>{t}</a:t></a:r></a:p>\
</c:rich></c:tx>\
<c:overlay val=\"0\"/>\
</c:title>"
    )
}

// ─────────────────────────────────────────────────────────────
// Plot area: dispatch to chart-type specific generators
// ─────────────────────────────────────────────────────────────

fn gen_xml_plot_area(chart: &ChartObject) -> String {
    match chart.chart_type {
        ChartType::Bar => gen_xml_bar_chart(chart),
        ChartType::Line => gen_xml_line_chart(chart),
        ChartType::Pie => gen_xml_pie_chart(chart),
        ChartType::Doughnut => gen_xml_doughnut_chart(chart),
        ChartType::Area => gen_xml_area_chart(chart),
        ChartType::Scatter => gen_xml_scatter_chart(chart),
        ChartType::Bubble | ChartType::Bubble3D => gen_xml_bubble_chart(chart),
        ChartType::StockHLC | ChartType::StockOHLC | ChartType::StockVHLC | ChartType::StockVOHLC => gen_xml_stock_chart(chart),
        ChartType::Surface | ChartType::SurfaceWireframe | ChartType::SurfaceTop | ChartType::SurfaceTopWireframe => gen_xml_surface_chart(chart),
        _ => gen_xml_bar_chart(chart), // fallback for unsupported types
    }
}

// ─────────────────────────────────────────────────────────────
// Bar / Column chart
// ─────────────────────────────────────────────────────────────

fn gen_xml_bar_chart(chart: &ChartObject) -> String {
    let opts = &chart.options;
    let bar_dir = match opts.bar_dir {
        BarDir::Column => "col",
        BarDir::Bar => "bar",
    };
    let grouping = match opts.bar_grouping {
        BarGrouping::Clustered => "clustered",
        BarGrouping::Stacked => "stacked",
        BarGrouping::PercentStacked => "percentStacked",
    };

    let mut s = format!(
        "<c:barChart>\
<c:barDir val=\"{bar_dir}\"/>\
<c:grouping val=\"{grouping}\"/>\
<c:varyColors val=\"0\"/>"
    );
    for (idx, ser) in chart.series.iter().enumerate() {
        s.push_str(&gen_xml_bar_series(ser, idx, opts.chart_colors.as_slice(), opts.show_value));
    }
    s.push_str(CHART_LEVEL_DLBLS);
    s.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
    s.push_str("</c:barChart>");
    // For horizontal bars the category axis is on the left, value axis on the bottom
    let (cat_pos, val_pos) = if opts.bar_dir == BarDir::Bar { ("l", "b") } else { ("b", "l") };
    s.push_str(&gen_xml_cat_ax(chart, 1, 2, cat_pos));
    s.push_str(&gen_xml_val_ax(chart, 2, 1, val_pos));
    s
}

fn gen_xml_bar_series(ser: &ChartSeries, idx: usize, chart_colors: &[String], show_val: bool) -> String {
    let color = ser.color.as_deref()
        .or_else(|| chart_colors.get(idx).map(|s| s.as_str()))
        .or_else(|| DEFAULT_CHART_COLORS.get(idx % DEFAULT_CHART_COLORS.len()).copied())
        .unwrap_or("4472C4");

    let mut s = format!(
        "<c:ser>\
<c:idx val=\"{idx}\"/>\
<c:order val=\"{idx}\"/>"
    );
    s.push_str(&gen_xml_ser_title(&ser.name));
    s.push_str(&format!(
        "<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>\
<a:ln><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></a:ln></c:spPr>"
    ));
    // dLbls must come before cat/val per CT_BarSer content model
    if show_val {
        s.push_str("<c:dLbls><c:numFmt formatCode=\"General\" sourceLinked=\"0\"/>\
<c:spPr><a:noFill/></c:spPr>\
<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/>\
<c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>");
    }
    s.push_str(&gen_xml_ser_cat(&ser.labels));
    s.push_str(&gen_xml_ser_val(&ser.values));
    s.push_str("</c:ser>");
    s
}

// ─────────────────────────────────────────────────────────────
// Line chart
// ─────────────────────────────────────────────────────────────

fn gen_xml_line_chart(chart: &ChartObject) -> String {
    let opts = &chart.options;
    let mut s = String::from(
        "<c:lineChart>\
<c:grouping val=\"standard\"/>\
<c:varyColors val=\"0\"/>"
    );
    for (idx, ser) in chart.series.iter().enumerate() {
        s.push_str(&gen_xml_line_series(ser, idx, opts.chart_colors.as_slice(), opts.show_value, opts.line_smooth, opts.show_markers));
    }
    s.push_str(CHART_LEVEL_DLBLS);
    s.push_str(&format!("<c:marker val=\"{}\"/>", if opts.show_markers { "1" } else { "0" }));
    // Chart-level smooth is always 0; per-series <c:smooth> handles actual smoothing
    s.push_str("<c:smooth val=\"0\"/>");
    s.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
    s.push_str("</c:lineChart>");
    s.push_str(&gen_xml_cat_ax(chart, 1, 2, "b"));
    s.push_str(&gen_xml_val_ax(chart, 2, 1, "l"));
    s
}

fn gen_xml_line_series(ser: &ChartSeries, idx: usize, chart_colors: &[String], show_val: bool, smooth: bool, markers: bool) -> String {
    let color = ser.color.as_deref()
        .or_else(|| chart_colors.get(idx).map(|s| s.as_str()))
        .or_else(|| DEFAULT_CHART_COLORS.get(idx % DEFAULT_CHART_COLORS.len()).copied())
        .unwrap_or("4472C4");

    let smooth_val = if smooth { "1" } else { "0" };
    let marker = if markers {
        format!("<c:marker><c:symbol val=\"circle\"/><c:size val=\"5\"/>\
<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></c:spPr></c:marker>")
    } else {
        "<c:marker><c:symbol val=\"none\"/></c:marker>".to_string()
    };

    let mut s = format!(
        "<c:ser>\
<c:idx val=\"{idx}\"/>\
<c:order val=\"{idx}\"/>"
    );
    s.push_str(&gen_xml_ser_title(&ser.name));
    s.push_str(&format!(
        "<c:spPr>\
<a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></a:ln></c:spPr>"
    ));
    s.push_str(&marker);
    // dLbls must come before cat/val per CT_LineSer content model
    if show_val {
        s.push_str("<c:dLbls><c:numFmt formatCode=\"General\" sourceLinked=\"0\"/>\
<c:spPr><a:noFill/></c:spPr>\
<c:showLegendKey val=\"0\"/><c:showVal val=\"1\"/><c:showCatName val=\"0\"/>\
<c:showSerName val=\"0\"/><c:showPercent val=\"0\"/><c:showBubbleSize val=\"0\"/></c:dLbls>");
    }
    s.push_str(&gen_xml_ser_cat(&ser.labels));
    s.push_str(&gen_xml_ser_val(&ser.values));
    s.push_str(&format!("<c:smooth val=\"{smooth_val}\"/>"));
    s.push_str("</c:ser>");
    s
}

// ─────────────────────────────────────────────────────────────
// Pie chart
// ─────────────────────────────────────────────────────────────

fn gen_xml_pie_chart(chart: &ChartObject) -> String {
    let opts = &chart.options;
    let first_angle = opts.first_slice_angle.unwrap_or(0);
    let mut s = format!(
        "<c:pieChart>\
<c:varyColors val=\"1\"/>"
    );
    // Pie charts typically have just one series
    if let Some(ser) = chart.series.first() {
        s.push_str(&gen_xml_pie_series(ser, 0, opts.chart_colors.as_slice(), opts.show_value));
    }
    s.push_str(CHART_LEVEL_DLBLS);
    s.push_str(&format!("<c:firstSliceAng val=\"{first_angle}\"/>"));
    s.push_str("</c:pieChart>");
    s
}

fn gen_xml_pie_series(ser: &ChartSeries, idx: usize, chart_colors: &[String], show_val: bool) -> String {
    let mut s = format!(
        "<c:ser>\
<c:idx val=\"{idx}\"/>\
<c:order val=\"{idx}\"/>"
    );
    s.push_str(&gen_xml_ser_title(&ser.name));
    // Individual data point colors
    for (pt_idx, _) in ser.values.iter().enumerate() {
        let color = chart_colors.get(pt_idx).map(|c| c.as_str())
            .or_else(|| DEFAULT_CHART_COLORS.get(pt_idx % DEFAULT_CHART_COLORS.len()).copied())
            .unwrap_or("4472C4");
        s.push_str(&format!(
            "<c:dPt><c:idx val=\"{pt_idx}\"/><c:bubble3D val=\"0\"/>\
<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>\
<a:ln><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill></a:ln></c:spPr></c:dPt>"
        ));
    }
    if show_val {
        s.push_str("<c:dLbls><c:numFmt formatCode=\"0%\" sourceLinked=\"0\"/>\
<c:spPr><a:noFill/></c:spPr>\
<c:showLegendKey val=\"0\"/><c:showVal val=\"0\"/><c:showCatName val=\"0\"/>\
<c:showSerName val=\"0\"/><c:showPercent val=\"1\"/><c:showBubbleSize val=\"0\"/></c:dLbls>");
    }
    s.push_str(&gen_xml_ser_cat(&ser.labels));
    s.push_str(&gen_xml_ser_val(&ser.values));
    s.push_str("</c:ser>");
    s
}

// ─────────────────────────────────────────────────────────────
// Doughnut chart
// ─────────────────────────────────────────────────────────────

fn gen_xml_doughnut_chart(chart: &ChartObject) -> String {
    let opts = &chart.options;
    let hole = opts.hole_size.unwrap_or(50);
    let first_angle = opts.first_slice_angle.unwrap_or(0);
    let mut s = format!(
        "<c:doughnutChart>\
<c:varyColors val=\"1\"/>"
    );
    if let Some(ser) = chart.series.first() {
        s.push_str(&gen_xml_pie_series(ser, 0, opts.chart_colors.as_slice(), opts.show_value));
    }
    s.push_str(CHART_LEVEL_DLBLS);
    s.push_str(&format!(
        "<c:firstSliceAng val=\"{first_angle}\"/>\
<c:holeSize val=\"{hole}\"/>"
    ));
    s.push_str("</c:doughnutChart>");
    s
}

// ─────────────────────────────────────────────────────────────
// Area chart
// ─────────────────────────────────────────────────────────────

fn gen_xml_area_chart(chart: &ChartObject) -> String {
    let opts = &chart.options;
    let mut s = String::from(
        "<c:areaChart>\
<c:grouping val=\"standard\"/>\
<c:varyColors val=\"0\"/>"
    );
    for (idx, ser) in chart.series.iter().enumerate() {
        let color = ser.color.as_deref()
            .or_else(|| opts.chart_colors.get(idx).map(|s| s.as_str()))
            .or_else(|| DEFAULT_CHART_COLORS.get(idx % DEFAULT_CHART_COLORS.len()).copied())
            .unwrap_or("4472C4");

        let mut ss = format!(
            "<c:ser>\
<c:idx val=\"{idx}\"/>\
<c:order val=\"{idx}\"/>"
        );
        ss.push_str(&gen_xml_ser_title(&ser.name));
        ss.push_str(&format!(
            "<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"><a:alpha val=\"80000\"/></a:srgbClr></a:solidFill>\
<a:ln><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></a:ln></c:spPr>"
        ));
        ss.push_str(&gen_xml_ser_cat(&ser.labels));
        ss.push_str(&gen_xml_ser_val(&ser.values));
        ss.push_str("</c:ser>");
        s.push_str(&ss);
    }
    s.push_str(CHART_LEVEL_DLBLS);
    s.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
    s.push_str("</c:areaChart>");
    s.push_str(&gen_xml_cat_ax(chart, 1, 2, "b"));
    s.push_str(&gen_xml_val_ax(chart, 2, 1, "l"));
    s
}

// ─────────────────────────────────────────────────────────────
// Scatter chart
// ─────────────────────────────────────────────────────────────

fn gen_xml_scatter_chart(chart: &ChartObject) -> String {
    let opts = &chart.options;
    let mut s = String::from(
        "<c:scatterChart>\
<c:scatterStyle val=\"lineMarker\"/>\
<c:varyColors val=\"0\"/>"
    );
    for (idx, ser) in chart.series.iter().enumerate() {
        let color = ser.color.as_deref()
            .or_else(|| opts.chart_colors.get(idx).map(|s| s.as_str()))
            .or_else(|| DEFAULT_CHART_COLORS.get(idx % DEFAULT_CHART_COLORS.len()).copied())
            .unwrap_or("4472C4");

        let mut ss = format!(
            "<c:ser>\
<c:idx val=\"{idx}\"/>\
<c:order val=\"{idx}\"/>"
        );
        ss.push_str(&gen_xml_ser_title(&ser.name));
        ss.push_str(&format!(
            "<c:spPr>\
<a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></a:ln></c:spPr>"
        ));
        ss.push_str(&format!(
            "<c:marker><c:symbol val=\"circle\"/><c:size val=\"5\"/>\
<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></c:spPr></c:marker>"
        ));
        // Scatter uses xVal and yVal (not cat/val)
        let x_vals: Vec<f64> = (0..ser.values.len()).map(|i| i as f64).collect();
        ss.push_str("<c:xVal>");
        ss.push_str(&gen_num_ref(&x_vals));
        ss.push_str("</c:xVal>");
        ss.push_str("<c:yVal>");
        ss.push_str(&gen_num_ref(&ser.values));
        ss.push_str("</c:yVal>");
        ss.push_str(&format!("<c:smooth val=\"{}\"/>", if opts.line_smooth { "1" } else { "0" }));
        ss.push_str("</c:ser>");
        s.push_str(&ss);
    }
    s.push_str(CHART_LEVEL_DLBLS);
    s.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
    s.push_str("</c:scatterChart>");
    s.push_str(&gen_xml_val_ax_raw(chart, 1, 2, "b", "x")); // x axis (bottom)
    s.push_str(&gen_xml_val_ax(chart, 2, 1, "l"));           // y axis (left)
    s
}

// ─────────────────────────────────────────────────────────────
// Axis XML helpers
// ─────────────────────────────────────────────────────────────

fn gen_xml_cat_ax(chart: &ChartObject, ax_id: u32, cross_ax_id: u32, pos: &str) -> String {
    let opts = &chart.options;
    let mut s = format!(
        "<c:catAx>\
<c:axId val=\"{ax_id}\"/>\
<c:scaling><c:orientation val=\"minMax\"/></c:scaling>\
<c:delete val=\"0\"/>\
<c:axPos val=\"{pos}\"/>"
    );
    if let Some(ref t) = opts.cat_axis_title {
        s.push_str(&gen_xml_axis_title(t));
    }
    s.push_str("<c:numFmt formatCode=\"General\" sourceLinked=\"0\"/>");
    s.push_str("<c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/>");
    s.push_str(&format!("<c:crossAx val=\"{cross_ax_id}\"/>"));
    s.push_str("</c:catAx>");
    s
}

fn gen_xml_val_ax(chart: &ChartObject, ax_id: u32, cross_ax_id: u32, pos: &str) -> String {
    let opts = &chart.options;
    let mut s = format!(
        "<c:valAx>\
<c:axId val=\"{ax_id}\"/>\
<c:scaling><c:orientation val=\"minMax\"/>"
    );
    if let Some(v) = opts.val_axis_max { s.push_str(&format!("<c:max val=\"{v}\"/>")); }
    if let Some(v) = opts.val_axis_min { s.push_str(&format!("<c:min val=\"{v}\"/>")); }
    s.push_str(&format!(
        "</c:scaling>\
<c:delete val=\"0\"/>\
<c:axPos val=\"{pos}\"/>"
    ));
    if opts.show_grid_lines {
        s.push_str("<c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val=\"D9D9D9\"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>");
    }
    if let Some(ref t) = opts.val_axis_title {
        s.push_str(&gen_xml_axis_title(t));
    }
    s.push_str("<c:numFmt formatCode=\"General\" sourceLinked=\"0\"/>");
    s.push_str("<c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/>");
    s.push_str(&format!("<c:crossAx val=\"{cross_ax_id}\"/>"));
    s.push_str("</c:valAx>");
    s
}

/// Same as val_ax but uses valAx element — for scatter chart x-axis
fn gen_xml_val_ax_raw(chart: &ChartObject, ax_id: u32, cross_ax_id: u32, pos: &str, _label: &str) -> String {
    let opts = &chart.options;
    let mut s = format!(
        "<c:valAx>\
<c:axId val=\"{ax_id}\"/>\
<c:scaling><c:orientation val=\"minMax\"/></c:scaling>\
<c:delete val=\"0\"/>\
<c:axPos val=\"{pos}\"/>"
    );
    if opts.show_grid_lines {
        s.push_str("<c:majorGridlines/>");
    }
    s.push_str("<c:numFmt formatCode=\"General\" sourceLinked=\"0\"/>");
    s.push_str("<c:majorTickMark val=\"out\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/>");
    s.push_str(&format!("<c:crossAx val=\"{cross_ax_id}\"/>"));
    s.push_str("</c:valAx>");
    s
}

fn gen_xml_axis_title(title: &str) -> String {
    let t = encode_xml_entities(title);
    format!(
        "<c:title>\
<c:tx><c:rich>\
<a:bodyPr rot=\"-5400000\"/><a:lstStyle/>\
<a:p><a:r><a:rPr lang=\"en-US\" dirty=\"0\"/><a:t>{t}</a:t></a:r></a:p>\
</c:rich></c:tx>\
<c:overlay val=\"0\"/>\
</c:title>"
    )
}

// ─────────────────────────────────────────────────────────────
// Series data helpers
// ─────────────────────────────────────────────────────────────

fn gen_xml_ser_title(name: &str) -> String {
    let n = encode_xml_entities(name);
    format!(
        "<c:tx><c:strRef>\
<c:f>Sheet1!$A$1</c:f>\
<c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>{n}</c:v></c:pt></c:strCache>\
</c:strRef></c:tx>"
    )
}

fn gen_xml_ser_cat(labels: &[String]) -> String {
    let count = labels.len();
    let mut pts = String::new();
    for (idx, lbl) in labels.iter().enumerate() {
        let v = encode_xml_entities(lbl);
        pts.push_str(&format!("<c:pt idx=\"{idx}\"><c:v>{v}</c:v></c:pt>"));
    }
    format!(
        "<c:cat><c:strRef>\
<c:f>Sheet1!$A$2:$A${}</c:f>\
<c:strCache><c:ptCount val=\"{count}\"/>{pts}</c:strCache>\
</c:strRef></c:cat>",
        count + 1
    )
}

fn gen_xml_ser_val(values: &[f64]) -> String {
    let count = values.len();
    let pts = gen_num_pts(values);
    format!(
        "<c:val><c:numRef>\
<c:f>Sheet1!$B$2:$B${}</c:f>\
<c:numCache>\
<c:formatCode>General</c:formatCode>\
<c:ptCount val=\"{count}\"/>{pts}\
</c:numCache>\
</c:numRef></c:val>",
        count + 1
    )
}

fn gen_num_ref(values: &[f64]) -> String {
    let count = values.len();
    let pts = gen_num_pts(values);
    format!(
        "<c:numRef>\
<c:f>Sheet1!$A$1:$A${count}</c:f>\
<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"{count}\"/>{pts}</c:numCache>\
</c:numRef>"
    )
}

fn gen_num_pts(values: &[f64]) -> String {
    let mut s = String::new();
    for (idx, v) in values.iter().enumerate() {
        s.push_str(&format!("<c:pt idx=\"{idx}\"><c:v>{v}</c:v></c:pt>"));
    }
    s
}

// ─────────────────────────────────────────────────────────────
// Bubble chart
// ─────────────────────────────────────────────────────────────

fn gen_xml_bubble_chart(chart: &ChartObject) -> String {
    let opts = &chart.options;
    let is_3d = chart.chart_type == ChartType::Bubble3D;

    let tag = if is_3d { "c:bubble3DChart" } else { "c:bubbleChart" };
    let mut s = format!("<{tag}>");
    s.push_str("<c:varyColors val=\"0\"/>");

    for (idx, ser) in chart.series.iter().enumerate() {
        let color = ser.color.as_deref()
            .or_else(|| opts.chart_colors.get(idx).map(|s| s.as_str()))
            .or_else(|| DEFAULT_CHART_COLORS.get(idx % DEFAULT_CHART_COLORS.len()).copied())
            .unwrap_or("4472C4");

        s.push_str(&format!("<c:ser><c:idx val=\"{idx}\"/><c:order val=\"{idx}\"/>"));
        s.push_str(&gen_xml_ser_title(&ser.name));
        s.push_str(&format!("<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></c:spPr>"));
        s.push_str("<c:invertIfNegative val=\"0\"/>");
        // X values (labels as numbers)
        let x_vals: Vec<f64> = (0..ser.values.len()).map(|i| i as f64).collect();
        s.push_str("<c:xVal>");
        s.push_str(&gen_num_ref(&x_vals));
        s.push_str("</c:xVal>");
        // Y values
        s.push_str("<c:yVal>");
        s.push_str(&gen_num_ref(&ser.values));
        s.push_str("</c:yVal>");
        // Bubble sizes
        if let Some(ref sizes) = ser.sizes {
            s.push_str("<c:bubbleSize>");
            s.push_str(&gen_num_ref(sizes));
            s.push_str("</c:bubbleSize>");
        }
        s.push_str("</c:ser>");
    }

    s.push_str(CHART_LEVEL_DLBLS);
    s.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
    s.push_str(&format!("</{tag}>"));

    // Value axes (X and Y) — bubble uses value axes for both
    s.push_str(&gen_xml_val_ax_raw(chart, 1, 2, "b", "x"));
    s.push_str(&gen_xml_val_ax(chart, 2, 1, "l"));
    s
}

// ─────────────────────────────────────────────────────────────
// Stock chart (HLC, OHLC, VHLC, VOHLC)
// ─────────────────────────────────────────────────────────────

fn gen_xml_stock_chart(chart: &ChartObject) -> String {
    let has_volume = matches!(chart.chart_type, ChartType::StockVHLC | ChartType::StockVOHLC);
    let has_open = matches!(chart.chart_type, ChartType::StockOHLC | ChartType::StockVOHLC);

    let mut s = String::new();
    let mut series_offset = 0;

    // Volume series as a bar chart (for VHLC/VOHLC)
    if has_volume && !chart.series.is_empty() {
        let vol = &chart.series[0];
        s.push_str("<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/><c:varyColors val=\"0\"/>");
        s.push_str(&gen_xml_bar_series(vol, 0, chart.options.chart_colors.as_slice(), false));
        s.push_str(CHART_LEVEL_DLBLS);
        s.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
        s.push_str("</c:barChart>");
        series_offset = 1;
    }

    // Stock chart
    s.push_str("<c:stockChart>");
    for (i, series) in chart.series.iter().skip(series_offset).enumerate() {
        let si = i + series_offset;
        s.push_str(&format!("<c:ser><c:idx val=\"{si}\"/><c:order val=\"{si}\"/>"));
        s.push_str(&gen_xml_ser_title(&series.name));
        s.push_str(&gen_xml_ser_cat(&series.labels));
        s.push_str(&gen_xml_ser_val(&series.values));
        s.push_str("</c:ser>");
    }
    s.push_str("<c:hiLowLines/>");
    if has_open {
        s.push_str("<c:upDownBars><c:gapWidth val=\"150\"/><c:upBars/><c:downBars/></c:upDownBars>");
    }
    s.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
    s.push_str("</c:stockChart>");

    // Axes
    s.push_str(&gen_xml_cat_ax(chart, 1, 2, "b"));
    s.push_str(&gen_xml_val_ax(chart, 2, 1, "l"));
    s
}

// ─────────────────────────────────────────────────────────────
// Surface chart
// ─────────────────────────────────────────────────────────────

fn gen_xml_surface_chart(chart: &ChartObject) -> String {
    let opts = &chart.options;
    let is_3d = matches!(chart.chart_type, ChartType::Surface | ChartType::SurfaceWireframe);
    let is_wireframe = matches!(chart.chart_type, ChartType::SurfaceWireframe | ChartType::SurfaceTopWireframe);

    let mut s = String::new();

    let tag = if is_3d { "c:surface3DChart" } else { "c:surfaceChart" };
    s.push_str(&format!("<{tag}>"));
    if is_wireframe {
        s.push_str("<c:wireframe val=\"1\"/>");
    }

    for (idx, ser) in chart.series.iter().enumerate() {
        let color = ser.color.as_deref()
            .or_else(|| opts.chart_colors.get(idx).map(|s| s.as_str()))
            .or_else(|| DEFAULT_CHART_COLORS.get(idx % DEFAULT_CHART_COLORS.len()).copied())
            .unwrap_or("4472C4");

        s.push_str(&format!("<c:ser><c:idx val=\"{idx}\"/><c:order val=\"{idx}\"/>"));
        s.push_str(&gen_xml_ser_title(&ser.name));
        s.push_str(&format!("<c:spPr><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill></c:spPr>"));
        s.push_str(&gen_xml_ser_cat(&ser.labels));
        s.push_str(&gen_xml_ser_val(&ser.values));
        s.push_str("</c:ser>");
    }

    // Band formats (for non-wireframe with multiple series)
    if !is_wireframe && chart.series.len() > 1 {
        s.push_str("<c:bandFmts>");
        for i in 0..chart.series.len() {
            s.push_str(&format!("<c:bandFmt><c:idx val=\"{i}\"/></c:bandFmt>"));
        }
        s.push_str("</c:bandFmts>");
    }

    s.push_str("<c:axId val=\"1\"/><c:axId val=\"2\"/>");
    if is_3d {
        s.push_str("<c:axId val=\"3\"/>"); // series axis for 3D
    }
    s.push_str(&format!("</{tag}>"));

    // Axes
    s.push_str(&gen_xml_cat_ax(chart, 1, 2, "b"));
    s.push_str(&gen_xml_val_ax(chart, 2, 1, "l"));
    if is_3d {
        s.push_str("<c:serAx><c:axId val=\"3\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:serAx>");
    }
    s
}

// ─────────────────────────────────────────────────────────────
// Chart relationship files
// ─────────────────────────────────────────────────────────────

/// Generates ppt/charts/_rels/chartN.xml.rels (typically empty — no external data source)
pub fn gen_xml_chart_rels() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>"
}
