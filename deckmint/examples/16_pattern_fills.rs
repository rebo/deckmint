//! 4x3 grid of shapes showing 12 different PatternType fills.

use deckmint::layout::{GridLayoutBuilder, GridTrack};
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, PatternFill, PatternType, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();

    // Dark background for contrast
    slide.set_background_color("#1A1A2E");

    slide.add_text(
        "Pattern Fills Gallery",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.15, 9.0, 0.55)
            .font_size(24.0)
            .bold()
            .color("#FFFFFF")
            .build(),
    );

    // 4 columns x 3 rows grid
    let grid = GridLayoutBuilder::new()
        .cols(vec![GridTrack::Fr(1.0); 4])
        .rows(vec![GridTrack::Fr(1.0); 3])
        .gap(0.15)
        .origin(0.4, 0.85)
        .container(9.2, 4.5)
        .build();

    // 12 pattern types with distinct foreground colors
    let patterns: [(PatternType, &str, &str); 12] = [
        (PatternType::Cross,     "Cross",      "#4472C4"),
        (PatternType::DnDiag,    "DnDiag",     "#ED7D31"),
        (PatternType::Horz,      "Horz",       "#70AD47"),
        (PatternType::Vert,      "Vert",       "#FFC000"),
        (PatternType::DotDmnd,   "DotDmnd",    "#5B9BD5"),
        (PatternType::LgCheck,   "LgCheck",    "#FF6384"),
        (PatternType::SmCheck,   "SmCheck",    "#36A2EB"),
        (PatternType::HorzBrick, "HorzBrick",  "#9966FF"),
        (PatternType::Wave,      "Wave",       "#FF9F40"),
        (PatternType::ZigZag,    "ZigZag",     "#4BC0C0"),
        (PatternType::Trellis,   "Trellis",    "#C9CBCF"),
        (PatternType::Sphere,    "Sphere",     "#E74C3C"),
    ];

    for (i, (pattern, label, fg_color)) in patterns.iter().enumerate() {
        let col = i % 4;
        let row = i / 4;
        let cell = grid.cell(col, row);

        // Shape with pattern fill
        slide.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .rect(cell.inset(0.05))
                .pattern_fill(PatternFill {
                    pattern: pattern.clone(),
                    fg_color: fg_color.trim_start_matches('#').into(),
                    bg_color: "FFFFFF".into(),
                })
                .rect_radius(0.08)
                .build(),
        );

        // Label centered in the cell
        slide.add_text(
            *label,
            TextOptionsBuilder::new()
                .rect(cell)
                .font_size(12.0)
                .bold()
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    pres.write_to_file("16_pattern_fills.pptx").unwrap();
    println!("Wrote 16_pattern_fills.pptx");
}
