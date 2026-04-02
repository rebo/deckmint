use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::table::{TableCell, TableOptionsBuilder};
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "deckmint Sample".to_string();
    pres.author = "deckmint".to_string();

    // Slide 1: text + shape
    let s1 = pres.add_slide();
    s1.add_text(
        "deckmint",
        TextOptionsBuilder::new()
            .x(0.5).y(0.5).w(9.0).h(1.5)
            .font_size(48.0).bold().align(AlignH::Center)
            .color("4472C4")
            .build(),
    );
    s1.add_text(
        "A Rust port of PptxGenJS",
        TextOptionsBuilder::new()
            .x(0.5).y(2.5).w(9.0).h(1.0)
            .font_size(24.0).align(AlignH::Center)
            .build(),
    );
    s1.add_shape(
        ShapeType::Rect,
        ShapeOptionsBuilder::new()
            .x(0.5).y(4.0).w(9.0).h(0.1)
            .fill_color("4472C4")
            .build(),
    );
    s1.add_notes("This is slide 1 speaker notes.");

    // Slide 2: table
    let s2 = pres.add_slide();
    s2.add_text(
        "Data Table",
        TextOptionsBuilder::new()
            .x(0.5).y(0.3).w(9.0).h(0.8)
            .font_size(28.0).bold()
            .build(),
    );
    let rows = vec![
        vec![TableCell::new("Name"), TableCell::new("Role"), TableCell::new("Score")],
        vec![TableCell::new("Alice"), TableCell::new("Engineer"), TableCell::new("98")],
        vec![TableCell::new("Bob"), TableCell::new("Designer"), TableCell::new("87")],
        vec![TableCell::new("Carol"), TableCell::new("PM"), TableCell::new("92")],
    ];
    s2.add_table(
        rows,
        TableOptionsBuilder::new()
            .x(0.5).y(1.5).w(9.0).h(3.0)
            .col_w(vec![3.0, 3.0, 3.0])
            .build(),
    );

    // Slide 3: dark background + white text
    let s3 = pres.add_slide();
    s3.set_background_color("1F3864");
    s3.add_text(
        "Dark Slide",
        TextOptionsBuilder::new()
            .x(0.5).y(2.0).w(9.0).h(1.5)
            .font_size(40.0).bold().align(AlignH::Center)
            .color("FFFFFF")
            .build(),
    );

    let out = "sample.pptx";
    pres.write_to_file(out).expect("write failed");
    println!("Wrote {out}");
}
