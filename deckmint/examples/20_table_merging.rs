//! Tables with colspan and rowspan — a weekly class timetable with merged cells.

use deckmint::objects::table::{TableCell, TableOptionsBuilder};
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, BorderProps, Presentation};

fn main() {
    let mut pres = Presentation::new();

    // ══════════════════════════════════════════════════════════
    // Slide 1: Weekly class timetable with merged time slots
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#F0F2F5");

        slide.add_text(
            "Weekly Class Schedule",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(26.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        let hdr_bg = "#1B2A4A";
        let hdr_fg = "#FFFFFF";
        let time_bg = "#E8EDF2";
        let math_bg = "#DCEEFB";
        let eng_bg = "#FDE8D0";
        let sci_bg = "#D4EDDA";
        let art_bg = "#F3E5F5";
        let pe_bg = "#FFF3CD";
        let lunch_bg = "#FFE0E0";
        let free_bg = "#F5F5F5";

        // Header row
        let header = vec![
            TableCell::new("Time").bold().fill(hdr_bg).color(hdr_fg).align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Monday").bold().fill(hdr_bg).color(hdr_fg).align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Tuesday").bold().fill(hdr_bg).color(hdr_fg).align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Wednesday").bold().fill(hdr_bg).color(hdr_fg).align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Thursday").bold().fill(hdr_bg).color(hdr_fg).align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Friday").bold().fill(hdr_bg).color(hdr_fg).align(AlignH::Center).valign(AlignV::Middle),
        ];

        // 8:00-9:00: Math spans Mon-Tue (colspan 2), English Wed-Thu (colspan 2), Art on Fri
        let row1 = vec![
            TableCell::new("8:00-9:00").bold().fill(time_bg).align(AlignH::Center).valign(AlignV::Middle).font_size(10.0),
            TableCell::new("Mathematics").fill(math_bg).color("#0D47A1").align(AlignH::Center).valign(AlignV::Middle).colspan(2),
            TableCell::merged(),
            TableCell::new("English").fill(eng_bg).color("#E65100").align(AlignH::Center).valign(AlignV::Middle).colspan(2),
            TableCell::merged(),
            TableCell::new("Art").fill(art_bg).color("#6A1B9A").align(AlignH::Center).valign(AlignV::Middle),
        ];

        // 9:00-10:00: Science spans Mon-Wed (colspan 3), PE spans Thu-Fri (colspan 2)
        let row2 = vec![
            TableCell::new("9:00-10:00").bold().fill(time_bg).align(AlignH::Center).valign(AlignV::Middle).font_size(10.0),
            TableCell::new("Science").fill(sci_bg).color("#1B5E20").align(AlignH::Center).valign(AlignV::Middle).colspan(3),
            TableCell::merged(),
            TableCell::merged(),
            TableCell::new("P.E.").fill(pe_bg).color("#F57F17").align(AlignH::Center).valign(AlignV::Middle).colspan(2),
            TableCell::merged(),
        ];

        // 10:00-10:30: Lunch break spans all 5 day columns
        let row3 = vec![
            TableCell::new("10:00-10:30").bold().fill(time_bg).align(AlignH::Center).valign(AlignV::Middle).font_size(10.0),
            TableCell::new("Morning Break").fill(lunch_bg).color("#C62828").bold().align(AlignH::Center).valign(AlignV::Middle).colspan(5),
            TableCell::merged(),
            TableCell::merged(),
            TableCell::merged(),
            TableCell::merged(),
        ];

        // 10:30-11:30: English Mon, Math Tue-Wed (colspan 2), Science Thu, Art Fri
        let row4 = vec![
            TableCell::new("10:30-11:30").bold().fill(time_bg).align(AlignH::Center).valign(AlignV::Middle).font_size(10.0),
            TableCell::new("English").fill(eng_bg).color("#E65100").align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Mathematics").fill(math_bg).color("#0D47A1").align(AlignH::Center).valign(AlignV::Middle).colspan(2),
            TableCell::merged(),
            TableCell::new("Science").fill(sci_bg).color("#1B5E20").align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Art").fill(art_bg).color("#6A1B9A").align(AlignH::Center).valign(AlignV::Middle),
        ];

        // 11:30-12:30: Free study period — individual cells
        let row5 = vec![
            TableCell::new("11:30-12:30").bold().fill(time_bg).align(AlignH::Center).valign(AlignV::Middle).font_size(10.0),
            TableCell::new("P.E.").fill(pe_bg).color("#F57F17").align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Art").fill(art_bg).color("#6A1B9A").align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Free Study").fill(free_bg).color("#757575").align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("English").fill(eng_bg).color("#E65100").align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Mathematics").fill(math_bg).color("#0D47A1").align(AlignH::Center).valign(AlignV::Middle),
        ];

        let rows = vec![header, row1, row2, row3, row4, row5];
        slide.add_table(
            rows,
            TableOptionsBuilder::new()
                .bounds(0.3, 1.0, 9.4, 4.3)
                .col_w(vec![1.2, 1.64, 1.64, 1.64, 1.64, 1.64])
                .font_size(11.0)
                .border(BorderProps::default())
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Project status matrix with rowspan
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#F0F2F5");

        slide.add_text(
            "Project Status Matrix",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(26.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        let hdr_bg = "#2C3E50";
        let hdr_fg = "#FFFFFF";
        let dept_bg = "#34495E";
        let dept_fg = "#ECF0F1";
        let done_bg = "#D5F5E3";
        let prog_bg = "#FEF9E7";
        let risk_bg = "#FADBD8";

        let header = vec![
            TableCell::new("Department").bold().fill(hdr_bg).color(hdr_fg).align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Project").bold().fill(hdr_bg).color(hdr_fg).align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Status").bold().fill(hdr_bg).color(hdr_fg).align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Owner").bold().fill(hdr_bg).color(hdr_fg).align(AlignH::Center).valign(AlignV::Middle),
        ];

        // Engineering spans 3 rows
        let eng_row1 = vec![
            TableCell::new("Engineering").bold().fill(dept_bg).color(dept_fg).align(AlignH::Center).valign(AlignV::Middle).rowspan(3),
            TableCell::new("API Redesign").align(AlignH::Left).valign(AlignV::Middle),
            TableCell::new("Complete").fill(done_bg).color("#1E8449").bold().align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Alice").align(AlignH::Center).valign(AlignV::Middle),
        ];
        let eng_row2 = vec![
            TableCell::merged(),
            TableCell::new("Cloud Migration").align(AlignH::Left).valign(AlignV::Middle),
            TableCell::new("In Progress").fill(prog_bg).color("#B7950B").bold().align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Bob").align(AlignH::Center).valign(AlignV::Middle),
        ];
        let eng_row3 = vec![
            TableCell::merged(),
            TableCell::new("Security Audit").align(AlignH::Left).valign(AlignV::Middle),
            TableCell::new("At Risk").fill(risk_bg).color("#C0392B").bold().align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Carol").align(AlignH::Center).valign(AlignV::Middle),
        ];

        // Design spans 2 rows
        let des_row1 = vec![
            TableCell::new("Design").bold().fill(dept_bg).color(dept_fg).align(AlignH::Center).valign(AlignV::Middle).rowspan(2),
            TableCell::new("Brand Refresh").align(AlignH::Left).valign(AlignV::Middle),
            TableCell::new("In Progress").fill(prog_bg).color("#B7950B").bold().align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Diana").align(AlignH::Center).valign(AlignV::Middle),
        ];
        let des_row2 = vec![
            TableCell::merged(),
            TableCell::new("Design System").align(AlignH::Left).valign(AlignV::Middle),
            TableCell::new("Complete").fill(done_bg).color("#1E8449").bold().align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Eve").align(AlignH::Center).valign(AlignV::Middle),
        ];

        // Marketing spans 2 rows
        let mkt_row1 = vec![
            TableCell::new("Marketing").bold().fill(dept_bg).color(dept_fg).align(AlignH::Center).valign(AlignV::Middle).rowspan(2),
            TableCell::new("Q2 Campaign").align(AlignH::Left).valign(AlignV::Middle),
            TableCell::new("Complete").fill(done_bg).color("#1E8449").bold().align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Frank").align(AlignH::Center).valign(AlignV::Middle),
        ];
        let mkt_row2 = vec![
            TableCell::merged(),
            TableCell::new("SEO Overhaul").align(AlignH::Left).valign(AlignV::Middle),
            TableCell::new("At Risk").fill(risk_bg).color("#C0392B").bold().align(AlignH::Center).valign(AlignV::Middle),
            TableCell::new("Grace").align(AlignH::Center).valign(AlignV::Middle),
        ];

        let rows = vec![header, eng_row1, eng_row2, eng_row3, des_row1, des_row2, mkt_row1, mkt_row2];
        slide.add_table(
            rows,
            TableOptionsBuilder::new()
                .bounds(0.5, 1.0, 9.0, 4.2)
                .col_w(vec![2.0, 2.8, 2.0, 2.2])
                .font_size(12.0)
                .border(BorderProps::default())
                .build(),
        );
    }

    pres.write_to_file("20_table_merging.pptx").unwrap();
    println!("Wrote 20_table_merging.pptx");
}
