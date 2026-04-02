//! Tables with styled headers, cell chaining, and per-cell formatting.

use deckmint::objects::table::{TableCell, TableOptionsBuilder};
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, BorderProps, Presentation};

fn main() {
    let mut pres = Presentation::new();

    let slide = pres.add_slide();
    slide.add_text(
        "Styled Table",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.3, 9.0, 0.8)
            .font_size(28.0).bold()
            .build(),
    );

    // Header row using chaining methods
    let header = vec![
        TableCell::new("Name").bold().fill("#2E4057").color("#FFFFFF"),
        TableCell::new("Department").bold().fill("#2E4057").color("#FFFFFF"),
        TableCell::new("Score").bold().fill("#2E4057").color("#FFFFFF").align(AlignH::Right),
    ];

    // Data rows
    let rows = vec![
        header,
        vec![
            TableCell::new("Alice"),
            TableCell::new("Engineering"),
            TableCell::new("98").align(AlignH::Right).color("#70AD47"),
        ],
        vec![
            TableCell::new("Bob"),
            TableCell::new("Design"),
            TableCell::new("87").align(AlignH::Right),
        ],
        vec![
            TableCell::new("Carol"),
            TableCell::new("Product").italic(),
            TableCell::new("92").align(AlignH::Right).color("#70AD47"),
        ],
    ];

    slide.add_table(
        rows,
        TableOptionsBuilder::new()
            .bounds(0.5, 1.5, 9.0, 3.5)
            .col_w(vec![3.0, 3.5, 2.5])
            .font_size(14.0)
            .border(BorderProps::default())
            .build(),
    );

    pres.write_to_file("04_tables.pptx").unwrap();
    println!("Wrote 04_tables.pptx");
}
