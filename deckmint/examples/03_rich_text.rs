//! Rich text with multiple styled runs, glow, outline, and tab stops.

use deckmint::objects::text::{TextOptionsBuilder, TextRunBuilder, TabStop};
use deckmint::{AlignH, GlowProps, Presentation};

fn main() {
    let mut pres = Presentation::new();

    let slide = pres.add_slide();
    slide.add_text(
        "Rich Text Demo",
        TextOptionsBuilder::new()
            .pos(0.5, 0.3).size(9.0, 0.8)
            .font_size(28.0).bold()
            .build(),
    );

    // Multi-run styled text
    slide.add_text_runs(
        vec![
            TextRunBuilder::new("Bold red ")
                .bold().color("FF0000").build(),
            TextRunBuilder::new("italic blue ")
                .italic().color("4472C4").build(),
            TextRunBuilder::new("with glow")
                .font_size(24.0)
                .glow(GlowProps::new(8.0, "FFC000", 0.7))
                .build(),
        ],
        TextOptionsBuilder::new()
            .pos(0.5, 1.5).size(9.0, 1.5)
            .font_size(20.0)
            .build(),
    );

    // Text with tab stops
    slide.add_text(
        "Left\tCenter\tRight",
        TextOptionsBuilder::new()
            .pos(0.5, 3.5).size(9.0, 1.0)
            .font_size(18.0)
            .tab_stops(vec![
                TabStop::new(3.0, "ctr"),
                TabStop::new(6.0, "r"),
            ])
            .build(),
    );

    // Subscript / superscript
    slide.add_text_runs(
        vec![
            TextRunBuilder::new("H").font_size(24.0).build(),
            TextRunBuilder::new("2").font_size(24.0).subscript().build(),
            TextRunBuilder::new("O  +  E=mc").font_size(24.0).build(),
            TextRunBuilder::new("2").font_size(24.0).superscript().build(),
        ],
        TextOptionsBuilder::new()
            .pos(0.5, 4.5).size(9.0, 1.0)
            .align(AlignH::Center)
            .build(),
    );

    pres.write_to_file("03_rich_text.pptx").unwrap();
    println!("Wrote 03_rich_text.pptx");
}
