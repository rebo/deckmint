//! Custom shapes via CustomGeomPoint: star, arrow, lightning bolt, speech bubble.

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, CustomGeomPoint as CGP, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();

    slide.set_background_color("#F0F0F0");

    slide.add_text(
        "Custom Geometry Shapes",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.15, 9.0, 0.55)
            .font_size(24.0)
            .bold()
            .color("#333333")
            .build(),
    );

    // ── Shape 1: 5-pointed star ─────────────────────────
    // Points computed for a regular 5-pointed star (outer r=0.5, inner r=0.2, centered at 0.5,0.5)
    let star = vec![
        CGP::MoveTo(0.50, 0.00),   // top
        CGP::LineTo(0.59, 0.35),
        CGP::LineTo(0.98, 0.35),   // right-upper tip
        CGP::LineTo(0.65, 0.57),
        CGP::LineTo(0.79, 0.91),   // right-lower tip
        CGP::LineTo(0.50, 0.70),
        CGP::LineTo(0.21, 0.91),   // left-lower tip
        CGP::LineTo(0.35, 0.57),
        CGP::LineTo(0.02, 0.35),   // left-upper tip
        CGP::LineTo(0.41, 0.35),
        CGP::Close,
    ];

    slide.add_shape(
        ShapeType::CustGeom,
        ShapeOptionsBuilder::new()
            .bounds(0.5, 1.0, 2.0, 2.0)
            .custom_geometry(star)
            .fill_color("#FFC107")
            .line_color("#F57F17")
            .line_width(1.5)
            .build(),
    );
    slide.add_text(
        "Star",
        TextOptionsBuilder::new()
            .bounds(0.5, 3.1, 2.0, 0.4)
            .font_size(14.0)
            .bold()
            .color("#333333")
            .align(AlignH::Center)
            .build(),
    );

    // ── Shape 2: Right-pointing arrow ───────────────────
    let arrow = vec![
        CGP::MoveTo(0.00, 0.30),
        CGP::LineTo(0.60, 0.30),
        CGP::LineTo(0.60, 0.00),
        CGP::LineTo(1.00, 0.50),   // arrow tip
        CGP::LineTo(0.60, 1.00),
        CGP::LineTo(0.60, 0.70),
        CGP::LineTo(0.00, 0.70),
        CGP::Close,
    ];

    slide.add_shape(
        ShapeType::CustGeom,
        ShapeOptionsBuilder::new()
            .bounds(3.0, 1.0, 2.0, 2.0)
            .custom_geometry(arrow)
            .fill_color("#4472C4")
            .line_color("#2F5597")
            .line_width(1.5)
            .build(),
    );
    slide.add_text(
        "Arrow",
        TextOptionsBuilder::new()
            .bounds(3.0, 3.1, 2.0, 0.4)
            .font_size(14.0)
            .bold()
            .color("#333333")
            .align(AlignH::Center)
            .build(),
    );

    // ── Shape 3: Lightning bolt ─────────────────────────
    let lightning = vec![
        CGP::MoveTo(0.55, 0.00),
        CGP::LineTo(0.20, 0.45),
        CGP::LineTo(0.45, 0.45),
        CGP::LineTo(0.30, 0.70),
        CGP::LineTo(0.50, 0.70),
        CGP::LineTo(0.35, 1.00),   // bottom tip
        CGP::LineTo(0.75, 0.55),
        CGP::LineTo(0.55, 0.55),
        CGP::LineTo(0.70, 0.30),
        CGP::LineTo(0.50, 0.30),
        CGP::LineTo(0.70, 0.00),
        CGP::Close,
    ];

    slide.add_shape(
        ShapeType::CustGeom,
        ShapeOptionsBuilder::new()
            .bounds(5.5, 1.0, 1.8, 2.0)
            .custom_geometry(lightning)
            .fill_color("#FF5722")
            .line_color("#BF360C")
            .line_width(1.5)
            .build(),
    );
    slide.add_text(
        "Lightning",
        TextOptionsBuilder::new()
            .bounds(5.5, 3.1, 1.8, 0.4)
            .font_size(14.0)
            .bold()
            .color("#333333")
            .align(AlignH::Center)
            .build(),
    );

    // ── Shape 4: Speech bubble (with Bezier curves) ─────
    let bubble = vec![
        // Start at top-left, trace rounded rectangle body using cubic Bezier corners
        CGP::MoveTo(0.15, 0.00),
        CGP::LineTo(0.85, 0.00),
        CGP::CubicBezTo(0.93, 0.00, 1.00, 0.07, 1.00, 0.15),  // top-right corner
        CGP::LineTo(1.00, 0.55),
        CGP::CubicBezTo(1.00, 0.63, 0.93, 0.70, 0.85, 0.70),  // bottom-right corner
        // Tail: down from the bottom edge then back up
        CGP::LineTo(0.40, 0.70),
        CGP::LineTo(0.20, 1.00),   // tail tip
        CGP::LineTo(0.30, 0.70),
        CGP::LineTo(0.15, 0.70),
        CGP::CubicBezTo(0.07, 0.70, 0.00, 0.63, 0.00, 0.55),  // bottom-left corner
        CGP::LineTo(0.00, 0.15),
        CGP::CubicBezTo(0.00, 0.07, 0.07, 0.00, 0.15, 0.00),  // top-left corner
        CGP::Close,
    ];

    slide.add_shape(
        ShapeType::CustGeom,
        ShapeOptionsBuilder::new()
            .bounds(7.8, 1.0, 2.0, 2.0)
            .custom_geometry(bubble)
            .fill_color("#E8F5E9")
            .line_color("#388E3C")
            .line_width(2.0)
            .build(),
    );
    slide.add_text(
        "Speech Bubble",
        TextOptionsBuilder::new()
            .bounds(7.8, 3.1, 2.0, 0.4)
            .font_size(14.0)
            .bold()
            .color("#333333")
            .align(AlignH::Center)
            .build(),
    );

    // Subtitle with technique note
    slide.add_text(
        "All shapes use CustomGeomPoint with normalized 0.0\u{2013}1.0 coordinates. \
         The speech bubble uses CubicBezTo for rounded corners.",
        TextOptionsBuilder::new()
            .bounds(0.5, 3.7, 9.0, 1.5)
            .font_size(12.0)
            .color("#666666")
            .align(AlignH::Center)
            .valign(AlignV::Top)
            .build(),
    );

    pres.write_to_file("19_custom_geometry.pptx").unwrap();
    println!("Wrote 19_custom_geometry.pptx");
}
