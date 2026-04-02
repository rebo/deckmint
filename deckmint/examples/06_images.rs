//! Images using the simplified `add_image_from` builder API.

use deckmint::objects::image::ImageOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{HyperlinkProps, Presentation, ShadowProps};

fn main() {
    let mut pres = Presentation::new();

    let slide = pres.add_slide();
    slide.add_text(
        "Image Examples",
        TextOptionsBuilder::new()
            .pos(0.5, 0.3).size(9.0, 0.8)
            .font_size(28.0).bold()
            .build(),
    );

    // 1×1 red PNG pixel (smallest valid PNG)
    let red_png: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D,
        0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00,
        0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
        0x00, 0x00, 0x03, 0x00, 0x01, 0x36, 0x28, 0x19, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    // Simple image placement
    slide.add_image_from(
        ImageOptionsBuilder::new()
            .bytes(red_png.clone(), "png")
            .pos(0.5, 1.5).size(3.0, 2.0)
    ).unwrap();

    // Image with shadow
    slide.add_image_from(
        ImageOptionsBuilder::new()
            .bytes(red_png.clone(), "png")
            .pos(4.0, 1.5).size(2.5, 2.0)
            .shadow(ShadowProps::outer().with_blur(8.0))
    ).unwrap();

    // Image with hyperlink and rotation
    slide.add_image_from(
        ImageOptionsBuilder::new()
            .bytes(red_png.clone(), "png")
            .pos(7.0, 1.5).size(2.0, 2.0)
            .rotate(15.0)
            .hyperlink(HyperlinkProps::url("https://example.com"))
    ).unwrap();

    // Image with rounded corners and transparency
    slide.add_image_from(
        ImageOptionsBuilder::new()
            .bytes(red_png, "png")
            .pos(2.0, 4.0).size(6.0, 1.5)
            .rounding()
            .transparency(40.0)
    ).unwrap();

    pres.write_to_file("06_images.pptx").unwrap();
    println!("Wrote 06_images.pptx");
}
