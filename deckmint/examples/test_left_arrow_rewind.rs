//! Minimal repro for left-arrow animation rewind.
//!
//! Generates a 2-slide deck where slide 1 has three sequential entrance
//! animations. Open in PowerPoint, run the slideshow, advance forward with
//! right-arrow / click, then press left-arrow to test backward navigation.

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AnimationEffect, Direction, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();

    let s = pres.add_slide();
    s.add_text(
        "Click 3 times, then press LEFT arrow",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.3, 9.0, 0.7)
            .font_size(22.0)
            .bold()
            .build(),
    );

    s.add_text(
        "1. Fade in",
        TextOptionsBuilder::new()
            .bounds(1.0, 1.5, 4.0, 0.6)
            .font_size(18.0)
            .animation(AnimationEffect::fade_in())
            .build(),
    );

    s.add_shape(
        ShapeType::Rect,
        ShapeOptionsBuilder::new()
            .bounds(1.0, 2.5, 3.0, 1.0)
            .fill_color("#4472C4")
            .animation(AnimationEffect::fly_in(Direction::Left))
            .build(),
    );

    s.add_shape(
        ShapeType::Ellipse,
        ShapeOptionsBuilder::new()
            .bounds(1.0, 4.0, 3.0, 1.0)
            .fill_color("#70AD47")
            .animation(AnimationEffect::zoom_in())
            .build(),
    );

    let s = pres.add_slide();
    s.add_text(
        "Slide 2 — left-arrow from here should step back through slide 1's animations",
        TextOptionsBuilder::new()
            .bounds(0.5, 1.0, 9.0, 2.0)
            .font_size(18.0)
            .build(),
    );

    pres.write_to_file("test_left_arrow_rewind.pptx").unwrap();
    println!("Wrote test_left_arrow_rewind.pptx");
}
