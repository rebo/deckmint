use crate::enums::EMU;
use crate::objects::table::{TableObject, TableCell, CellTextRun};
use crate::presentation::Presentation;
use crate::types::{BorderType, PatternFill};
use crate::utils::{encode_xml_entities, gen_xml_color_selection_str, inch_to_emu, val_to_pts};

/// Generate the full `<p:graphicFrame>` XML for a table
pub fn gen_xml_table(
    tbl: &TableObject,
    obj_id: usize,
    pres: &Presentation,
) -> String {
    let layout = &pres.layout;
    let x = tbl.options.position.x.as_ref().map(|c| c.to_emu(layout.width)).unwrap_or(EMU);
    let y = tbl.options.position.y.as_ref().map(|c| c.to_emu(layout.height)).unwrap_or(EMU);
    let cx = tbl.options.position.w.as_ref().map(|c| c.to_emu(layout.width)).unwrap_or(EMU);
    let cy = tbl.options.position.h.as_ref().map(|c| c.to_emu(layout.height)).unwrap_or(EMU);

    let obj_name = tbl.options.object_name.as_deref().unwrap_or("Table 1");
    let frame_id = obj_id;

    // Count columns from first row (accounting for colspan)
    let col_count = count_columns(&tbl.rows);

    let tbl_pr = if let Some(ref style_id) = tbl.options.table_style_id {
        format!("<a:tblPr><a:tableStyleId>{style_id}</a:tableStyleId></a:tblPr>")
    } else {
        "<a:tblPr/>".to_string()
    };

    let mut s = String::new();
    s.push_str(&format!(
        "<p:graphicFrame>\
<p:nvGraphicFramePr>\
<p:cNvPr id=\"{frame_id}\" name=\"{obj_name}\"/>\
<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr>\
<p:nvPr><p:extLst><p:ext uri=\"{{D42A27DB-BD31-4B8C-83A1-F6EECF244321}}\">\
<p14:modId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"1579011935\"/>\
</p:ext></p:extLst></p:nvPr>\
</p:nvGraphicFramePr>\
<p:xfrm><a:off x=\"{x}\" y=\"{y}\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></p:xfrm>\
<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\">\
<a:tbl>{tbl_pr}"
    ));

    // Column widths
    s.push_str("<a:tblGrid>");
    if let Some(ref col_widths) = tbl.options.col_w {
        for i in 0..col_count {
            let w = col_widths.get(i).copied().map(inch_to_emu)
                .unwrap_or_else(|| cx / col_count as i64);
            s.push_str(&format!("<a:gridCol w=\"{w}\"/>"));
        }
    } else {
        let w = if col_count > 0 { cx / col_count as i64 } else { cx };
        for _ in 0..col_count {
            s.push_str(&format!("<a:gridCol w=\"{w}\"/>"));
        }
    }
    s.push_str("</a:tblGrid>");

    // Rows
    for (row_idx, row) in tbl.rows.iter().enumerate() {
        let row_h = tbl.options.row_h
            .as_ref()
            .and_then(|rh| rh.get(row_idx).copied())
            .map(inch_to_emu)
            .unwrap_or_else(|| inch_to_emu(0.3));
        s.push_str(&format!("<a:tr h=\"{row_h}\">"));
        for cell in row {
            s.push_str(&gen_xml_table_cell(cell, tbl, col_count, cx));
        }
        s.push_str("</a:tr>");
    }

    s.push_str("</a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
    s
}

fn count_columns(rows: &[Vec<TableCell>]) -> usize {
    rows.first()
        .map(|row| {
            row.iter()
                .map(|cell| cell.options.colspan.unwrap_or(1) as usize)
                .sum()
        })
        .unwrap_or(0)
}

fn gen_xml_table_cell(
    cell: &TableCell,
    tbl: &TableObject,
    _col_count: usize,
    _cx: i64,
) -> String {
    let mut s = String::new();

    // Merged cell placeholder
    if cell.is_merged {
        s.push_str("<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang=\"en-US\" dirty=\"0\"/></a:p></a:txBody><a:tcPr/></a:tc>");
        return s;
    }

    let mut tc_pr_attrs = String::new();
    if let Some(cs) = cell.options.colspan {
        if cs > 1 { tc_pr_attrs.push_str(&format!(" gridSpan=\"{cs}\"")); }
    }
    if let Some(rs) = cell.options.rowspan {
        if rs > 1 { tc_pr_attrs.push_str(&format!(" rowSpan=\"{rs}\"")); }
    }

    s.push_str(&format!("<a:tc{tc_pr_attrs}>"));

    // txBody
    s.push_str("<a:txBody><a:bodyPr/><a:lstStyle/>");

    if cell.text.is_empty() {
        s.push_str("<a:p><a:endParaRPr lang=\"en-US\" dirty=\"0\"/></a:p>");
    } else {
        // Group text runs by line breaks
        let mut current_para: Vec<&CellTextRun> = Vec::new();
        let mut paras: Vec<Vec<&CellTextRun>> = Vec::new();

        for run in &cell.text {
            current_para.push(run);
            if run.break_line {
                paras.push(current_para);
                current_para = Vec::new();
            }
        }
        if !current_para.is_empty() {
            paras.push(current_para);
        }

        for para in paras {
            s.push_str("<a:p>");
            // Paragraph properties
            let align = cell.options.align.as_ref()
                .or(tbl.options.align.as_ref());
            if let Some(a) = align {
                s.push_str(&format!("<a:pPr algn=\"{}\"><a:buNone/></a:pPr>", a.as_ooxml()));
            } else {
                s.push_str("<a:pPr indent=\"0\" marL=\"0\"><a:buNone/></a:pPr>");
            }

            for run in para {
                s.push_str("<a:r>");
                let mut rpr_attrs = String::from("<a:rPr lang=\"en-US\"");

                let fs = run.font_size.or(cell.options.font_size).or(tbl.options.font_size);
                if let Some(f) = fs { rpr_attrs.push_str(&format!(" sz=\"{}\"", (f * 100.0).round() as i64)); }

                let bold = run.bold.or(cell.options.bold).or(tbl.options.bold);
                if let Some(b) = bold { rpr_attrs.push_str(&format!(" b=\"{}\"", if b { 1 } else { 0 })); }

                let italic = run.italic.or(cell.options.italic).or(tbl.options.italic);
                if let Some(i) = italic { rpr_attrs.push_str(&format!(" i=\"{}\"", if i { 1 } else { 0 })); }

                if let Some(true) = run.underline.or(cell.options.underline) {
                    rpr_attrs.push_str(" u=\"sng\"");
                }

                rpr_attrs.push_str(" dirty=\"0\"");

                let mut rpr_children = String::new();

                let color = run.color.as_deref()
                    .or(cell.options.color.as_deref())
                    .or(tbl.options.color.as_deref());
                if let Some(c) = color {
                    rpr_children.push_str(&gen_xml_color_selection_str(c, None));
                }

                let ff = run.font_face.as_deref().or(cell.options.font_face.as_deref()).or(tbl.options.font_face.as_deref());
                if let Some(f) = ff {
                    rpr_children.push_str(&format!(
                        "<a:latin typeface=\"{f}\" pitchFamily=\"34\" charset=\"0\"/>"
                    ));
                }

                // Cell hyperlink (applied to each run)
                if let Some(ref hl) = cell.options.hyperlink {
                    let tooltip = hl.tooltip.as_deref().unwrap_or("");
                    let tooltip_attr = if tooltip.is_empty() { String::new() } else { format!(" tooltip=\"{}\"", encode_xml_entities(tooltip)) };
                    if let Some(ref nav) = hl.action {
                        rpr_children.push_str(&format!("<a:hlinkClick r:id=\"\" action=\"{}\"{tooltip_attr}/>", nav.as_ppaction()));
                    } else if hl.r_id > 0 {
                        if hl.slide.is_some() {
                            rpr_children.push_str(&format!("<a:hlinkClick r:id=\"rId{}\" action=\"ppaction://hlinksldjump\"{tooltip_attr}/>", hl.r_id));
                        } else if hl.url.is_some() {
                            rpr_children.push_str(&format!("<a:hlinkClick r:id=\"rId{}\"{tooltip_attr}/>", hl.r_id));
                        }
                    }
                }

                // Self-closing rPr when no children
                if rpr_children.is_empty() {
                    s.push_str(&rpr_attrs);
                    s.push_str("/>");
                } else {
                    s.push_str(&rpr_attrs);
                    s.push('>');
                    s.push_str(&rpr_children);
                    s.push_str("</a:rPr>");
                }
                s.push_str(&format!("<a:t>{}</a:t>", encode_xml_entities(&run.text)));
                s.push_str("</a:r>");
            }

            s.push_str("<a:endParaRPr lang=\"en-US\" dirty=\"0\"/>");
            s.push_str("</a:p>");
        }
    }

    s.push_str("</a:txBody>");

    // tcPr (cell properties: fill, borders, margins, valign)
    s.push_str("<a:tcPr");
    // Margins
    let margin = cell.options.margin.as_ref().or(tbl.options.margin.as_ref());
    if let Some(m) = margin {
        s.push_str(&format!(
            " marL=\"{}\" marR=\"{}\" marT=\"{}\" marB=\"{}\"",
            val_to_pts(m.left()),
            val_to_pts(m.right()),
            val_to_pts(m.top()),
            val_to_pts(m.bottom()),
        ));
    } else {
        s.push_str(&format!(
            " marL=\"{}\" marR=\"{}\" marT=\"{}\" marB=\"{}\"",
            val_to_pts(0.1),
            val_to_pts(0.1),
            val_to_pts(0.05),
            val_to_pts(0.05),
        ));
    }
    // Vertical alignment
    if let Some(ref va) = cell.options.valign.as_ref().or(tbl.options.valign.as_ref()) {
        s.push_str(&format!(" anchor=\"{}\"", va.as_ooxml()));
    }
    s.push_str(">");

    // Cell fill — pattern > gradient > solid > table default
    let cell_grad = cell.options.gradient_fill.as_ref();
    let tbl_grad  = tbl.options.gradient_fill.as_ref();
    let grad = cell_grad.or(tbl_grad);
    if let Some(pf) = cell.options.pattern_fill.as_ref() {
        s.push_str(&gen_xml_pattern_fill(pf));
    } else if let Some(g) = grad {
        s.push_str(&crate::utils::gen_xml_grad_fill(g));
    } else {
        let fill = cell.options.fill.as_deref().or(tbl.options.fill.as_deref());
        if let Some(f) = fill {
            s.push_str(&gen_xml_color_selection_str(f, None));
        } else {
            s.push_str("<a:noFill/>");
        }
    }

    // Borders — use table-level default or cell-level override
    let def_border = tbl.options.border.as_ref();
    let cell_borders = cell.options.border.as_ref();

    for side in &["lnL", "lnR", "lnT", "lnB"] {
        let border = cell_borders.and_then(|cb| match *side {
            "lnL" => cb.left.as_ref(),
            "lnR" => cb.right.as_ref(),
            "lnT" => cb.top.as_ref(),
            "lnB" => cb.bottom.as_ref(),
            _ => None,
        }).or(def_border);

        if let Some(b) = border {
            let w = val_to_pts(b.pt);
            s.push_str(&format!("<a:{side} w=\"{w}\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">"));
            match b.border_type {
                BorderType::None => s.push_str("<a:noFill/>"),
                _ => {
                    if let Some(ref color) = b.color {
                        s.push_str(&gen_xml_color_selection_str(color, None));
                    }
                }
            }
            s.push_str(&format!("</a:{side}>"));
        }
    }

    s.push_str("</a:tcPr>");
    s.push_str("</a:tc>");
    s
}

fn gen_xml_pattern_fill(pf: &PatternFill) -> String {
    let prst = pf.pattern.as_str();
    let fg = pf.fg_color.trim_start_matches('#').to_uppercase();
    let bg = pf.bg_color.trim_start_matches('#').to_uppercase();
    format!(
        "<a:pattFill prst=\"{prst}\">\
<a:fgClr><a:srgbClr val=\"{fg}\"/></a:fgClr>\
<a:bgClr><a:srgbClr val=\"{bg}\"/></a:bgClr>\
</a:pattFill>"
    )
}
