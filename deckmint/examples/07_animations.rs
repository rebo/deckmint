//! Click-triggered entrance, exit, and emphasis animations.

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AnimationEffect, Direction, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();

    // Slide 1: Entrance animations
    let s = pres.add_slide();
    s.add_text("Entrance Animations", TextOptionsBuilder::new()
        .bounds(0.5, 0.3, 9.0, 0.7).font_size(24.0).bold().build());

    s.add_text("Click → I fade in", TextOptionsBuilder::new()
        .bounds(0.5, 1.5, 4.0, 0.8).font_size(18.0)
        .animation(AnimationEffect::fade_in())
        .build());

    s.add_text("Click → I fly from left", TextOptionsBuilder::new()
        .bounds(0.5, 2.5, 4.0, 0.8).font_size(18.0)
        .animation(AnimationEffect::fly_in(Direction::Left))
        .build());

    s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
        .bounds(5.5, 1.5, 3.0, 1.5)
        .fill_color("#4472C4")
        .animation(AnimationEffect::zoom_in())
        .build());

    s.add_shape(ShapeType::Ellipse, ShapeOptionsBuilder::new()
        .bounds(5.5, 3.5, 3.0, 1.5)
        .fill_color("#70AD47")
        .animation(AnimationEffect::bounce_in())
        .build());

    // Slide 2: Emphasis animations
    let s = pres.add_slide();
    s.add_text("Emphasis Animations", TextOptionsBuilder::new()
        .bounds(0.5, 0.3, 9.0, 0.7).font_size(24.0).bold().build());

    s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
        .bounds(0.5, 1.5, 2.5, 2.0)
        .fill_color("#4472C4")
        .animation(AnimationEffect::spin(360.0))
        .build());

    s.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
        .bounds(3.5, 1.5, 2.5, 2.0)
        .fill_color("#ED7D31")
        .animation(AnimationEffect::pulse())
        .build());

    s.add_shape(ShapeType::Ellipse, ShapeOptionsBuilder::new()
        .bounds(6.5, 1.5, 2.5, 2.0)
        .fill_color("#70AD47")
        .animation(AnimationEffect::grow_shrink(1.5))
        .build());

    // Slide 3: Exit animations
    let s = pres.add_slide();
    s.add_text("Exit Animations", TextOptionsBuilder::new()
        .bounds(0.5, 0.3, 9.0, 0.7).font_size(24.0).bold().build());

    s.add_text("Click → I disappear", TextOptionsBuilder::new()
        .bounds(0.5, 1.5, 4.0, 0.8).font_size(18.0)
        .animation(AnimationEffect::disappear())
        .build());

    s.add_text("Click → I fly out right", TextOptionsBuilder::new()
        .bounds(0.5, 2.5, 4.0, 0.8).font_size(18.0)
        .animation(AnimationEffect::fly_out(Direction::Right))
        .build());

    s.add_text("Click → I fade out", TextOptionsBuilder::new()
        .bounds(0.5, 3.5, 4.0, 0.8).font_size(18.0)
        .animation(AnimationEffect::fade_out())
        .build());

    pres.write_to_file("07_animations.pptx").unwrap();
    println!("Wrote 07_animations.pptx");
}
