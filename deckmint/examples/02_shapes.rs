//! Shapes with fills, gradients, shadows, and hyperlinks.

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{
    AlignH, GradientFill, HyperlinkProps, Presentation, ShadowProps, ShapeType,
};

fn main() {
    let mut pres = Presentation::new();

    let slide = pres.add_slide();
    slide.add_text(
        "Shape Gallery",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.3, 9.0, 0.8)
            .font_size(28.0).bold()
            .build(),
    );

    // Solid-fill rectangle with a drop shadow
    slide.add_shape(
        ShapeType::Rect,
        ShapeOptionsBuilder::new()
            .bounds(0.5, 1.5, 2.5, 2.0)
            .fill_color("#4472C4")
            .shadow(ShadowProps::outer().with_blur(6.0).with_offset(4.0))
            .build(),
    );

    // Gradient-fill rounded rectangle
    slide.add_shape(
        ShapeType::RoundRect,
        ShapeOptionsBuilder::new()
            .bounds(3.5, 1.5, 2.5, 2.0)
            .gradient_fill(GradientFill::two_color(90.0, "#4472C4", "#70AD47"))
            .rect_radius(0.2)
            .build(),
    );

    // Ellipse with hyperlink
    slide.add_shape(
        ShapeType::Ellipse,
        ShapeOptionsBuilder::new()
            .bounds(6.5, 1.5, 2.5, 2.0)
            .fill_color("#ED7D31")
            .hyperlink(HyperlinkProps::url("https://github.com").with_tooltip("GitHub"))
            .build(),
    );

    // Rotated triangle
    slide.add_shape(
        ShapeType::Triangle,
        ShapeOptionsBuilder::new()
            .bounds(2.0, 4.0, 2.0, 2.0)
            .fill_color("#A9D18E")
            .rotate(45.0)
            .build(),
    );

    // Inner shadow diamond
    slide.add_shape(
        ShapeType::Diamond,
        ShapeOptionsBuilder::new()
            .bounds(5.5, 4.0, 2.0, 2.0)
            .fill_color("#FFC000")
            .shadow(ShadowProps::inner().with_color("#000000").with_blur(8.0))
            .build(),
    );

    pres.write_to_file("02_shapes.pptx").unwrap();
    println!("Wrote 02_shapes.pptx");
}
