//! 6-slide quarterly report — corporate gray/blue palette with master slide,
//! KPI cards, column chart, styled table, line chart, and numbered action items.

use deckmint::layout::{GridLayoutBuilder, GridTrack};
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::table::{TableCell, TableOptionsBuilder};
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{
    AlignH, AlignV, ChartOptionsBuilder, ChartSeries, ChartType, LegendPos, Presentation,
    ShapeType, SlideMasterDef,
};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Q4 2025 Quarterly Report".to_string();

    // ── Colors ────────────────────────────────────────────────
    let dark_gray = "#1B2A4A";
    let mid_gray = "#2C3E50";
    let light_bg = "#F4F6F8";
    let accent = "#4472C4";
    let green = "#70AD47";
    let orange = "#ED7D31";
    let red = "#E74C3C";
    let white = "#FFFFFF";
    let muted = "#7F8C9B";

    // ── Master slide: thin header + footer ────────────────────
    {
        let master_grid = GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(1.0)])
            .rows(vec![
                GridTrack::Inches(0.5),
                GridTrack::Fr(1.0),
                GridTrack::Inches(0.3),
            ])
            .build();

        let header = master_grid.cell(0, 0);
        let footer = master_grid.cell(0, 2);

        let mut master = SlideMasterDef::new("Report");
        master.set_background_color(light_bg);

        // Header bar
        master.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .rect(header)
                .fill_color(dark_gray)
                .build(),
        );
        master.add_text(
            "ACME Corp  |  Q4 2025 Report",
            TextOptionsBuilder::new()
                .rect(header.inset_xy(0.3, 0.0))
                .font_size(11.0)
                .bold()
                .color(white)
                .valign(AlignV::Middle)
                .build(),
        );

        // Footer bar
        master.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .rect(footer)
                .fill_color("#E0E4E8")
                .build(),
        );
        master.add_text(
            "Confidential  |  For Internal Use Only",
            TextOptionsBuilder::new()
                .rect(footer.inset_xy(0.3, 0.0))
                .font_size(8.0)
                .color(muted)
                .valign(AlignV::Middle)
                .align(AlignH::Right)
                .build(),
        );

        pres.define_master(master);
    }

    // ══════════════════════════════════════════════════════════
    // Slide 1: Cover — "Q4 2025 Report"
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color(dark_gray);

        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 2.0, 10.0, 0.04)
                .fill_color(accent)
                .build(),
        );

        s.add_text(
            "Q4 2025",
            TextOptionsBuilder::new()
                .bounds(1.0, 1.0, 8.0, 1.0)
                .font_size(48.0)
                .bold()
                .color(white)
                .align(AlignH::Center)
                .valign(AlignV::Bottom)
                .build(),
        );

        s.add_text(
            "Quarterly Business Report",
            TextOptionsBuilder::new()
                .bounds(1.0, 2.2, 8.0, 0.8)
                .font_size(24.0)
                .color("#90A4AE")
                .align(AlignH::Center)
                .valign(AlignV::Top)
                .build(),
        );

        s.add_text(
            "ACME Corporation",
            TextOptionsBuilder::new()
                .bounds(1.0, 3.5, 8.0, 0.5)
                .font_size(16.0)
                .bold()
                .color(accent)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        s.add_text(
            "Prepared by Finance & Strategy  |  January 2026",
            TextOptionsBuilder::new()
                .bounds(1.0, 4.2, 8.0, 0.5)
                .font_size(11.0)
                .color(muted)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Executive Summary — 4 KPI cards
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "Executive Summary",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.6, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color(dark_gray)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(4, 1, 0.2)
            .origin(0.5, 1.4)
            .container(9.0, 1.8)
            .build();

        let kpis = [
            ("$14.2M", "Revenue", "+18% QoQ", green),
            ("$3.8M", "Net Income", "+12% QoQ", green),
            ("1,247", "New Customers", "+24% QoQ", accent),
            ("94.2%", "Retention", "-0.3%", orange),
        ];

        for (i, (value, label, delta, color)) in kpis.iter().enumerate() {
            let cell = grid.cell(i, 0);

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color(white)
                    .line_color("#D0D5DD")
                    .line_width(1.0)
                    .rect_radius(0.06)
                    .build(),
            );

            // Accent top bar
            s.add_shape(
                ShapeType::Rect,
                ShapeOptionsBuilder::new()
                    .bounds(cell.x, cell.y, cell.w, 0.04)
                    .fill_color(*color)
                    .build(),
            );

            let inner = cell.inset(0.12);
            s.add_text(
                *value,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y, inner.w, 0.7)
                    .font_size(28.0)
                    .bold()
                    .color(dark_gray)
                    .align(AlignH::Center)
                    .valign(AlignV::Bottom)
                    .build(),
            );
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y + 0.7, inner.w, 0.35)
                    .font_size(11.0)
                    .color(muted)
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
            s.add_text(
                *delta,
                TextOptionsBuilder::new()
                    .bounds(inner.x, inner.y + 1.0, inner.w, 0.3)
                    .font_size(11.0)
                    .bold()
                    .color(*color)
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }

        // Summary paragraph
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(0.5, 3.5, 9.0, 1.5)
                .fill_color(white)
                .line_color("#D0D5DD")
                .line_width(1.0)
                .rect_radius(0.06)
                .build(),
        );
        s.add_text(
            "Q4 marked our strongest quarter with revenue exceeding targets by 8%. Customer acquisition accelerated driven by the new enterprise tier launch in October. Retention dipped slightly due to seasonal churn in the SMB segment, but enterprise retention remains at 98.5%.",
            TextOptionsBuilder::new()
                .bounds(0.8, 3.7, 8.4, 1.1)
                .font_size(12.0)
                .color(mid_gray)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Revenue — column chart
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "Revenue Breakdown",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.6, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color(dark_gray)
                .build(),
        );

        let quarters = vec!["Q1", "Q2", "Q3", "Q4"];
        let series = vec![
            ChartSeries::new("Subscriptions", quarters.clone(), vec![6.2, 7.1, 8.0, 9.4]),
            ChartSeries::new("Services", quarters.clone(), vec![2.1, 2.4, 2.6, 3.0]),
            ChartSeries::new("Licensing", quarters.clone(), vec![1.0, 1.2, 1.5, 1.8]),
        ];

        s.add_chart(
            ChartType::Bar,
            series,
            ChartOptionsBuilder::new()
                .bounds(0.5, 1.2, 9.0, 4.0)
                .title("Revenue by Stream ($M)")
                .show_value()
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(vec![accent, green, orange])
                .val_axis_title("Revenue ($M)")
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Regional — styled table
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "Regional Performance",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.6, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color(dark_gray)
                .build(),
        );

        let hdr_bg = dark_gray;
        let alt_bg = "#EDF1F5";

        let header = vec![
            TableCell::new("Region").bold().fill(hdr_bg).color(white).align(AlignH::Center),
            TableCell::new("Revenue").bold().fill(hdr_bg).color(white).align(AlignH::Center),
            TableCell::new("Growth").bold().fill(hdr_bg).color(white).align(AlignH::Center),
            TableCell::new("Customers").bold().fill(hdr_bg).color(white).align(AlignH::Center),
            TableCell::new("NPS").bold().fill(hdr_bg).color(white).align(AlignH::Center),
        ];

        let rows = vec![
            header,
            vec![
                TableCell::new("North America").bold(),
                TableCell::new("$6.8M").align(AlignH::Right),
                TableCell::new("+22%").color(green).bold().align(AlignH::Center),
                TableCell::new("612").align(AlignH::Center),
                TableCell::new("72").align(AlignH::Center),
            ],
            vec![
                TableCell::new("Europe").bold().fill(alt_bg),
                TableCell::new("$3.9M").align(AlignH::Right).fill(alt_bg),
                TableCell::new("+15%").color(green).bold().align(AlignH::Center).fill(alt_bg),
                TableCell::new("384").align(AlignH::Center).fill(alt_bg),
                TableCell::new("68").align(AlignH::Center).fill(alt_bg),
            ],
            vec![
                TableCell::new("Asia Pacific").bold(),
                TableCell::new("$2.4M").align(AlignH::Right),
                TableCell::new("+31%").color(green).bold().align(AlignH::Center),
                TableCell::new("178").align(AlignH::Center),
                TableCell::new("65").align(AlignH::Center),
            ],
            vec![
                TableCell::new("Latin America").bold().fill(alt_bg),
                TableCell::new("$1.1M").align(AlignH::Right).fill(alt_bg),
                TableCell::new("-3%").color(red).bold().align(AlignH::Center).fill(alt_bg),
                TableCell::new("73").align(AlignH::Center).fill(alt_bg),
                TableCell::new("58").align(AlignH::Center).fill(alt_bg),
            ],
            vec![
                TableCell::new("Total").bold().fill(mid_gray).color(white),
                TableCell::new("$14.2M").bold().align(AlignH::Right).fill(mid_gray).color(white),
                TableCell::new("+18%").bold().align(AlignH::Center).fill(mid_gray).color(green),
                TableCell::new("1,247").bold().align(AlignH::Center).fill(mid_gray).color(white),
                TableCell::new("67").bold().align(AlignH::Center).fill(mid_gray).color(white),
            ],
        ];

        s.add_table(
            rows,
            TableOptionsBuilder::new()
                .bounds(0.5, 1.3, 9.0, 3.6)
                .col_w(vec![2.2, 1.8, 1.5, 1.8, 1.7])
                .font_size(12.0)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 5: Trends — line chart
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "12-Month Trends",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.6, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color(dark_gray)
                .build(),
        );

        let months = vec![
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ];

        let series = vec![
            ChartSeries::new(
                "MRR ($K)",
                months.clone(),
                vec![920.0, 955.0, 1010.0, 1045.0, 1080.0, 1120.0, 1155.0, 1200.0, 1250.0, 1310.0, 1380.0, 1420.0],
            ),
            ChartSeries::new(
                "New Customers",
                months,
                vec![85.0, 92.0, 98.0, 105.0, 112.0, 108.0, 115.0, 122.0, 130.0, 145.0, 138.0, 142.0],
            ),
        ];

        s.add_chart(
            ChartType::Line,
            series,
            ChartOptionsBuilder::new()
                .bounds(0.5, 1.2, 9.0, 4.0)
                .title("Monthly Recurring Revenue & Customer Growth")
                .show_legend(true)
                .legend_pos(LegendPos::Bottom)
                .chart_colors(vec![accent, green])
                .line_smooth()
                .cat_axis_title("Month")
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 6: Action Items — numbered list
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();

        s.add_text(
            "Q1 2026 Action Items",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.6, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color(dark_gray)
                .build(),
        );

        let items = [
            ("Launch enterprise self-serve portal", "Product", green),
            ("Expand APAC sales team by 5 reps", "Sales", accent),
            ("Reduce SMB churn to <5% monthly", "Customer Success", orange),
            ("Ship AI-powered analytics module", "Engineering", accent),
            ("Achieve SOC 2 Type II certification", "Security", red),
            ("Close 3 strategic partnerships", "BD", green),
        ];

        let list_area = deckmint::layout::CellRect {
            x: 0.5, y: 1.3, w: 9.0, h: 3.8,
        };
        let rows = deckmint::layout::split_v(&list_area, items.len(), 0.1);

        for (i, (task, owner, color)) in items.iter().enumerate() {
            let row = rows[i];

            // Row background
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(row)
                    .fill_color(white)
                    .line_color("#E0E4E8")
                    .line_width(1.0)
                    .rect_radius(0.04)
                    .build(),
            );

            // Number circle
            let num_size = 0.35;
            let num_x = row.x + 0.15;
            let num_y = row.y + (row.h - num_size) / 2.0;
            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(num_x, num_y, num_size, num_size)
                    .fill_color(*color)
                    .build(),
            );
            s.add_text(
                &format!("{}", i + 1),
                TextOptionsBuilder::new()
                    .bounds(num_x, num_y, num_size, num_size)
                    .font_size(12.0)
                    .bold()
                    .color(white)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Task text
            s.add_text(
                *task,
                TextOptionsBuilder::new()
                    .bounds(row.x + 0.65, row.y, 5.5, row.h)
                    .font_size(13.0)
                    .bold()
                    .color(dark_gray)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Owner tag
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(row.x + 6.5, row.y + (row.h - 0.3) / 2.0, 2.2, 0.3)
                    .fill_color(*color)
                    .rect_radius(0.04)
                    .build(),
            );
            s.add_text(
                *owner,
                TextOptionsBuilder::new()
                    .bounds(row.x + 6.5, row.y + (row.h - 0.3) / 2.0, 2.2, 0.3)
                    .font_size(9.0)
                    .bold()
                    .color(white)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }
    }

    pres.write_to_file("31_report.pptx").unwrap();
    println!("Wrote 31_report.pptx");
}
