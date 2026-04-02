//! Images using the simplified `add_image_from` builder API.

use deckmint::objects::image::ImageOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{HyperlinkProps, Presentation, ShadowProps};

fn main() {
    let mut pres = Presentation::new();

    let tiger_png = std::fs::read(concat!(env!("CARGO_MANIFEST_DIR"), "/examples/tiger.png"))
        .expect("tiger.png not found — place it in deckmint/examples/");

    let slide = pres.add_slide();
    slide.add_text(
        "Image Examples",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.3, 9.0, 0.8)
            .font_size(28.0).bold()
            .build(),
    );

    // Simple image placement
    slide.add_image_from(
        ImageOptionsBuilder::new()
            .bytes(tiger_png.clone(), "png")
            .bounds(0.5, 1.5, 3.0, 2.0)
    ).unwrap();

    // Image with shadow
    slide.add_image_from(
        ImageOptionsBuilder::new()
            .bytes(tiger_png.clone(), "png")
            .bounds(4.0, 1.5, 2.5, 2.0)
            .shadow(ShadowProps::outer().with_blur(8.0))
    ).unwrap();

    // Image with hyperlink and rotation
    slide.add_image_from(
        ImageOptionsBuilder::new()
            .bytes(tiger_png.clone(), "png")
            .bounds(7.0, 1.5, 2.0, 2.0)
            .rotate(15.0)
            .hyperlink(HyperlinkProps::url("https://example.com"))
    ).unwrap();

    // Image with rounded corners and transparency
    slide.add_image_from(
        ImageOptionsBuilder::new()
            .bytes(tiger_png, "png")
            .bounds(2.0, 4.0, 6.0, 1.5)
            .rounding()
            .transparency(40.0)
    ).unwrap();

    pres.write_to_file("06_images.pptx").unwrap();
    println!("Wrote 06_images.pptx");
}
