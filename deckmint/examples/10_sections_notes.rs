//! Sections, speaker notes, slide numbers, and hidden slides.

use deckmint::objects::text::TextOptionsBuilder;
use deckmint::types::SlideNumberProps;
use deckmint::{AlignH, Presentation};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Sections & Notes Demo".to_string();

    // Section A
    pres.add_section("Introduction");

    let s = pres.add_slide();
    s.add_text(
        "Welcome",
        TextOptionsBuilder::new()
            .pos(0.5, 2.0).size(9.0, 1.5)
            .font_size(44.0).bold().align(AlignH::Center)
            .color("4472C4")
            .build(),
    );
    s.add_notes("Welcome the audience. Introduce the topic.");
    s.set_slide_number(SlideNumberProps {
        x: Some(deckmint::Coord::Inches(9.0)),
        y: Some(deckmint::Coord::Inches(5.2)),
        font_size: Some(10.0),
        color: Some("999999".to_string()),
        ..Default::default()
    });

    // Section B
    pres.add_section("Content");

    let s = pres.add_slide();
    s.add_text(
        "Slide 2 — Content Section",
        TextOptionsBuilder::new()
            .pos(0.5, 2.0).size(9.0, 1.5)
            .font_size(32.0).align(AlignH::Center)
            .build(),
    );
    s.add_notes("Main content goes here.\nRemember to cover all three points.");

    let s = pres.add_slide();
    s.add_text(
        "Slide 3 — Also in Content",
        TextOptionsBuilder::new()
            .pos(0.5, 2.0).size(9.0, 1.5)
            .font_size(32.0).align(AlignH::Center)
            .build(),
    );

    // A hidden backup slide
    let s = pres.add_slide();
    s.add_text(
        "Hidden Backup Slide",
        TextOptionsBuilder::new()
            .pos(0.5, 2.0).size(9.0, 1.5)
            .font_size(28.0).align(AlignH::Center)
            .color("999999")
            .build(),
    );
    s.hide();
    s.add_notes("This slide is hidden and won't show during the presentation.");

    pres.write_to_file("10_sections_notes.pptx").unwrap();
    println!("Wrote 10_sections_notes.pptx");
}
