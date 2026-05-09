//! Auto-firing cascade that left-arrow can rewind.
//!
//! Mimics the MCQ-deck pattern: first animation uses `.after_previous().delay(N)`
//! so the cascade fires automatically on slide enter (no click), and subsequent
//! animations chain after it. The whole step should still be rewindable via
//! the left arrow during slideshow.
//!
//! Manual test:
//!   1. Open in PowerPoint, no repair prompt expected.
//!   2. Start slideshow on slide 1 → cascade auto-fires after 1s.
//!   3. Wait for cascade to finish.
//!   4. Press LEFT → effects undo one by one back to slide start.
//!   5. Press LEFT again → goes to slide 0 (or previous slide).

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AnimationEffect, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();

    // Slide 0: anchor for left-arrow navigation tests.
    let s = pres.add_slide();
    s.add_text(
        "Slide 0 — advance to slide 1 to start the cascade.",
        TextOptionsBuilder::new()
            .bounds(0.5, 1.0, 9.0, 2.0)
            .font_size(20.0)
            .build(),
    );

    // Slide 1: auto-firing cascade.
    let s = pres.add_slide();
    s.add_text(
        "Auto-fires after 1s. LEFT should rewind through the effects.",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.3, 9.0, 0.7)
            .font_size(22.0)
            .bold()
            .build(),
    );

    // Effect 1: leading .after_previous().delay(1000) → auto step, leader fires
    // 1000 ms after slide enter.
    s.add_text(
        "Effect 1 (auto, +1000ms)",
        TextOptionsBuilder::new()
            .bounds(1.0, 1.5, 4.0, 0.6)
            .font_size(18.0)
            .animation(AnimationEffect::fade_in().after_previous().delay(1000))
            .build(),
    );

    // Effect 2 & 3: chain after.
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

    pres.write_to_file("test_auto_rewind.pptx").unwrap();
    println!("Wrote test_auto_rewind.pptx");
}
