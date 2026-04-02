//! Slide backgrounds — solid colours, transparency, and images.

use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, Presentation};

fn main() {
    let mut pres = Presentation::new();

    // 1×1 blue PNG pixel for background image
    let blue_png: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
        0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00,
        0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0x60, 0x60, 0xF8, 0x0F,
        0x00, 0x00, 0x01, 0x01, 0x00, 0x05, 0x18, 0xD8, 0x4D, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    // Slide 1: Dark solid background
    let s = pres.add_slide();
    s.set_background_color("1B2838");
    s.add_text(
        "Dark Background",
        TextOptionsBuilder::new()
            .pos(0.5, 2.0).size(9.0, 1.5)
            .font_size(40.0).bold()
            .align(AlignH::Center).valign(AlignV::Middle)
            .color("FFFFFF")
            .build(),
    );

    // Slide 2: Coloured background with transparency
    let s = pres.add_slide();
    s.set_background_color_transparency("4472C4", 30.0);
    s.add_text(
        "Semi-transparent Blue (30%)",
        TextOptionsBuilder::new()
            .pos(0.5, 2.0).size(9.0, 1.5)
            .font_size(32.0).bold()
            .align(AlignH::Center).valign(AlignV::Middle)
            .build(),
    );

    // Slide 3: Image background
    let s = pres.add_slide();
    s.set_background_image(blue_png, "png");
    s.add_text(
        "Image Background",
        TextOptionsBuilder::new()
            .pos(0.5, 2.0).size(9.0, 1.5)
            .font_size(40.0).bold()
            .align(AlignH::Center).valign(AlignV::Middle)
            .color("FFFFFF")
            .build(),
    );

    // Slide 4: Warm gradient feel via solid colour
    let s = pres.add_slide();
    s.set_background_color("FFF2CC");
    s.add_text(
        "Warm Background",
        TextOptionsBuilder::new()
            .pos(0.5, 2.0).size(9.0, 1.5)
            .font_size(36.0).bold()
            .align(AlignH::Center).valign(AlignV::Middle)
            .color("BF8F00")
            .build(),
    );

    pres.write_to_file("08_backgrounds.pptx").unwrap();
    println!("Wrote 08_backgrounds.pptx");
}
