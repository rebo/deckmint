//! Minimal "Hello World" presentation — one slide, one text box.

use deckmint::Presentation;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::AlignH;

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Hello World".to_string();

    let slide = pres.add_slide();
    slide.add_text(
        "Hello, World!",
        TextOptionsBuilder::new()
            .bounds(1.0, 2.0, 8.0, 2.0)
            .font_size(44.0)
            .bold()
            .align(AlignH::Center)
            .color("#4472C4")
            .build(),
    );

    pres.write_to_file("01_hello_world.pptx").unwrap();
    println!("Wrote 01_hello_world.pptx");
}
