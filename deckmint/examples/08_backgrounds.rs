//! Slide backgrounds — solid colours, transparency, and images.

use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, Presentation};

fn main() {
    let mut pres = Presentation::new();

    let tiger_png = std::fs::read(concat!(env!("CARGO_MANIFEST_DIR"), "/examples/tiger.png"))
        .expect("tiger.png not found — place it in deckmint/examples/");

    // Slide 1: Dark solid background
    let s = pres.add_slide();
    s.set_background_color("#1B2838");
    s.add_text(
        "Dark Background",
        TextOptionsBuilder::new()
            .bounds(0.5, 2.0, 9.0, 1.5)
            .font_size(40.0).bold()
            .align(AlignH::Center).valign(AlignV::Middle)
            .color("#FFFFFF")
            .build(),
    );

    // Slide 2: Coloured background with transparency
    let s = pres.add_slide();
    s.set_background_color_transparency("#4472C4", 30.0);
    s.add_text(
        "Semi-transparent Blue (30%)",
        TextOptionsBuilder::new()
            .bounds(0.5, 2.0, 9.0, 1.5)
            .font_size(32.0).bold()
            .align(AlignH::Center).valign(AlignV::Middle)
            .build(),
    );

    // Slide 3: Image background
    let s = pres.add_slide();
    s.set_background_image(tiger_png, "png");
    s.add_text(
        "Image Background",
        TextOptionsBuilder::new()
            .bounds(0.5, 2.0, 9.0, 1.5)
            .font_size(40.0).bold()
            .align(AlignH::Center).valign(AlignV::Middle)
            .color("#FFFFFF")
            .build(),
    );

    // Slide 4: Warm gradient feel via solid colour
    let s = pres.add_slide();
    s.set_background_color("#FFF2CC");
    s.add_text(
        "Warm Background",
        TextOptionsBuilder::new()
            .bounds(0.5, 2.0, 9.0, 1.5)
            .font_size(36.0).bold()
            .align(AlignH::Center).valign(AlignV::Middle)
            .color("#BF8F00")
            .build(),
    );

    pres.write_to_file("08_backgrounds.pptx").unwrap();
    println!("Wrote 08_backgrounds.pptx");
}
