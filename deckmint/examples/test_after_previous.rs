//! Repro for timed-cascade animations.
//!
//! Three sequential entrance animations on slide 1:
//!   1. fade_in           (OnClick)
//!   2. fade_in().after_previous().delay(500)
//!   3. fade_in().after_previous().delay(500)
//!
//! Expected (matching hand-authored timing.pptx):
//!   - One outer <p:par> wraps all three effects (one click triggers the cascade)
//!   - All three share the same grpId
//!   - Effects 2/3 use nodeType="afterEffect"
//!   - Left-arrow rewinds the cascade to slide start
//!
//! Open in PowerPoint: there should be NO repair prompt. Run slideshow,
//! click once on slide 1 — all three should appear in sequence. Then
//! press LEFT to test rewind.

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AnimationEffect, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();

    let s = pres.add_slide();
    s.add_text(
        "1 click → 3 effects cascade. Then LEFT to rewind.",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.3, 9.0, 0.7)
            .font_size(22.0)
            .bold()
            .build(),
    );

    s.add_text(
        "Effect 1 (OnClick)",
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
            .animation(AnimationEffect::fade_in().after_previous().delay(500))
            .build(),
    );

    s.add_shape(
        ShapeType::Ellipse,
        ShapeOptionsBuilder::new()
            .bounds(1.0, 4.0, 3.0, 1.0)
            .fill_color("#70AD47")
            .animation(AnimationEffect::fade_in().after_previous().delay(500))
            .build(),
    );

    let s = pres.add_slide();
    s.add_text(
        "Slide 2 — left-arrow from here should rewind to slide 1.",
        TextOptionsBuilder::new()
            .bounds(0.5, 1.0, 9.0, 2.0)
            .font_size(18.0)
            .build(),
    );

    pres.write_to_file("test_after_previous.pptx").unwrap();
    println!("Wrote test_after_previous.pptx");
}
