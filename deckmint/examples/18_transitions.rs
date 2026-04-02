//! 10 slides each with a different transition type.

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::types::{TransitionProps, TransitionSpeed, TransitionType};
use deckmint::{AlignH, AlignV, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();

    let transitions: [(TransitionType, &str, &str); 10] = [
        (TransitionType::Fade,    "Fade",    "#4472C4"),
        (TransitionType::Push,    "Push",    "#ED7D31"),
        (TransitionType::Wipe,    "Wipe",    "#70AD47"),
        (TransitionType::Split,   "Split",   "#FFC000"),
        (TransitionType::Cover,   "Cover",   "#5B9BD5"),
        (TransitionType::Zoom,    "Zoom",    "#9B59B6"),
        (TransitionType::Flash,   "Flash",   "#E74C3C"),
        (TransitionType::Diamond, "Diamond", "#1ABC9C"),
        (TransitionType::Blinds,  "Blinds",  "#F39C12"),
        (TransitionType::Checker, "Checker", "#2C3E50"),
    ];

    for (i, (transition_type, name, color)) in transitions.iter().enumerate() {
        let slide = pres.add_slide();
        slide.set_background_color("#0D1117");

        // Set this slide's transition
        slide.set_transition(TransitionProps {
            transition_type: transition_type.clone(),
            speed: Some(TransitionSpeed::Medium),
            direction: None,
            duration_ms: None,
            advance_after_ms: Some(3000),
            advance_on_click: true,
        });

        // Colored accent bar at top
        slide.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .bounds(0.0, 0.0, 10.0, 0.08)
                .fill_color(*color)
                .build(),
        );

        // Slide number indicator
        slide.add_text(
            &format!("{} / {}", i + 1, transitions.len()),
            TextOptionsBuilder::new()
                .bounds(0.5, 0.3, 9.0, 0.5)
                .font_size(14.0)
                .color("#666666")
                .align(AlignH::Right)
                .build(),
        );

        // Large transition name
        slide.add_text(
            *name,
            TextOptionsBuilder::new()
                .bounds(0.5, 1.5, 9.0, 2.0)
                .font_size(60.0)
                .bold()
                .color(*color)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Subtitle
        slide.add_text(
            "This slide uses the transition shown above.\nAuto-advances after 3 seconds.",
            TextOptionsBuilder::new()
                .bounds(0.5, 3.6, 9.0, 1.0)
                .font_size(16.0)
                .color("#888888")
                .align(AlignH::Center)
                .valign(AlignV::Top)
                .build(),
        );

        // Decorative shape echoing the transition color
        slide.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(3.5, 4.7, 3.0, 0.4)
                .fill_color(*color)
                .rect_radius(0.2)
                .build(),
        );
    }

    pres.write_to_file("18_transitions.pptx").unwrap();
    println!("Wrote 18_transitions.pptx");
}
