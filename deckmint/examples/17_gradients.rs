//! Gradient fills: multi-stop linear, radial, and gradient text boxes.

use deckmint::layout::{GridLayoutBuilder, GridTrack};
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, GradientFill, GradientStop, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();

    // ══════════════════════════════════════════════════════════
    // Slide 1: Multi-stop linear gradients at various angles
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#F5F5F5");

        slide.add_text(
            "Multi-Stop Linear Gradients",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(24.0)
                .bold()
                .color("#333333")
                .build(),
        );

        let grid = GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(1.0); 3])
            .rows(vec![GridTrack::Fr(1.0); 2])
            .gap(0.2)
            .origin(0.5, 1.0)
            .container(9.0, 4.3)
            .build();

        // Gradient definitions: (label, angle, stops)
        let gradients: [(&str, f64, Vec<GradientStop>); 6] = [
            (
                "Sunset (0\u{00b0})",
                0.0,
                vec![
                    GradientStop::new("#FF512F", 0.0),
                    GradientStop::new("#F09819", 50.0),
                    GradientStop::new("#DD2476", 100.0),
                ],
            ),
            (
                "Ocean (90\u{00b0})",
                90.0,
                vec![
                    GradientStop::new("#2193B0", 0.0),
                    GradientStop::new("#6DD5ED", 50.0),
                    GradientStop::new("#FFFFFF", 100.0),
                ],
            ),
            (
                "Forest (45\u{00b0})",
                45.0,
                vec![
                    GradientStop::new("#134E5E", 0.0),
                    GradientStop::new("#71B280", 50.0),
                    GradientStop::new("#C6FFDD", 100.0),
                ],
            ),
            (
                "Neon (135\u{00b0})",
                135.0,
                vec![
                    GradientStop::new("#A855F7", 0.0),
                    GradientStop::new("#EC4899", 40.0),
                    GradientStop::new("#F97316", 70.0),
                    GradientStop::new("#FACC15", 100.0),
                ],
            ),
            (
                "Arctic (180\u{00b0})",
                180.0,
                vec![
                    GradientStop::new("#E0EAFC", 0.0),
                    GradientStop::new("#CFDEF3", 50.0),
                    GradientStop::new("#4A90D9", 100.0),
                ],
            ),
            (
                "Fire (270\u{00b0})",
                270.0,
                vec![
                    GradientStop::new("#F83600", 0.0),
                    GradientStop::new("#F9D423", 50.0),
                    GradientStop::new("#FE8C00", 100.0),
                ],
            ),
        ];

        for (i, (label, angle, stops)) in gradients.iter().enumerate() {
            let col = i % 3;
            let row = i / 3;
            let cell = grid.cell(col, row);

            slide.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell.inset(0.05))
                    .gradient_fill(GradientFill::linear(*angle, stops.clone()))
                    .rect_radius(0.12)
                    .build(),
            );
            slide.add_text(
                *label,
                TextOptionsBuilder::new()
                    .rect(cell)
                    .font_size(14.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Bottom)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Radial gradients
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#1A1A2E");

        slide.add_text(
            "Radial Gradients",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(24.0)
                .bold()
                .color("#FFFFFF")
                .build(),
        );

        let grid = GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(1.0); 3])
            .rows(vec![GridTrack::Fr(1.0); 2])
            .gap(0.2)
            .origin(0.5, 1.0)
            .container(9.0, 4.3)
            .build();

        // Radial gradient definitions
        let radials: [(&str, Vec<GradientStop>); 6] = [
            (
                "Hot Core",
                vec![
                    GradientStop::new("#FFFFFF", 0.0),
                    GradientStop::new("#FFC107", 40.0),
                    GradientStop::new("#FF5722", 100.0),
                ],
            ),
            (
                "Cool Orb",
                vec![
                    GradientStop::new("#E0F7FA", 0.0),
                    GradientStop::new("#4DD0E1", 50.0),
                    GradientStop::new("#006064", 100.0),
                ],
            ),
            (
                "Purple Glow",
                vec![
                    GradientStop::new("#E1BEE7", 0.0),
                    GradientStop::new("#9C27B0", 60.0),
                    GradientStop::new("#4A148C", 100.0),
                ],
            ),
            (
                "Simple: Red-Blue",
                vec![
                    GradientStop::new("#FF0000", 0.0),
                    GradientStop::new("#0000FF", 100.0),
                ],
            ),
            (
                "Emerald Sphere",
                vec![
                    GradientStop::new("#A5D6A7", 0.0),
                    GradientStop::new("#388E3C", 50.0),
                    GradientStop::new("#1B5E20", 100.0),
                ],
            ),
            (
                "Gold Disc",
                vec![
                    GradientStop::new("#FFF9C4", 0.0),
                    GradientStop::new("#FFD54F", 40.0),
                    GradientStop::new("#F57F17", 100.0),
                ],
            ),
        ];

        for (i, (label, stops)) in radials.iter().enumerate() {
            let col = i % 3;
            let row = i / 3;
            let cell = grid.cell(col, row);

            slide.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .rect(cell.inset(0.1))
                    .gradient_fill(GradientFill::radial(stops.clone()))
                    .build(),
            );
            slide.add_text(
                *label,
                TextOptionsBuilder::new()
                    .rect(cell)
                    .font_size(12.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Bottom)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Gradient text boxes with convenience constructors
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#0D1117");

        slide.add_text(
            "Gradient Text Boxes",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(24.0)
                .bold()
                .color("#FFFFFF")
                .build(),
        );

        // Row 1: two_color convenience
        slide.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(0.5, 1.1, 4.2, 1.6)
                .gradient_fill(GradientFill::two_color(0.0, "#667EEA", "#764BA2"))
                .rect_radius(0.1)
                .build(),
        );
        slide.add_text(
            "two_color(0, blue, purple)\nSimple left-to-right",
            TextOptionsBuilder::new()
                .bounds(0.5, 1.1, 4.2, 1.6)
                .font_size(14.0)
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        slide.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(5.3, 1.1, 4.2, 1.6)
                .gradient_fill(GradientFill::two_color(90.0, "#11998E", "#38EF7D"))
                .rect_radius(0.1)
                .build(),
        );
        slide.add_text(
            "two_color(90, teal, green)\nTop-to-bottom",
            TextOptionsBuilder::new()
                .bounds(5.3, 1.1, 4.2, 1.6)
                .font_size(14.0)
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Row 2: radial_two_color convenience
        slide.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(0.5, 3.0, 4.2, 1.6)
                .gradient_fill(GradientFill::radial_two_color("#FFFFFF", "#E74C3C"))
                .rect_radius(0.1)
                .build(),
        );
        slide.add_text(
            "radial_two_color\nwhite center, red edge",
            TextOptionsBuilder::new()
                .bounds(0.5, 3.0, 4.2, 1.6)
                .font_size(14.0)
                .color("#333333")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Multi-stop with transparency
        slide.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(5.3, 3.0, 4.2, 1.6)
                .gradient_fill(GradientFill::linear(
                    45.0,
                    vec![
                        GradientStop::new("#FF0080", 0.0),
                        GradientStop::new("#7928CA", 50.0).with_transparency(30.0),
                        GradientStop::new("#FF0080", 100.0),
                    ],
                ))
                .rect_radius(0.1)
                .build(),
        );
        slide.add_text(
            "Multi-stop + transparency\nat 50% midpoint",
            TextOptionsBuilder::new()
                .bounds(5.3, 3.0, 4.2, 1.6)
                .font_size(14.0)
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    pres.write_to_file("17_gradients.pptx").unwrap();
    println!("Wrote 17_gradients.pptx");
}
