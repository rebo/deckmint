//! Showcase of 40+ ShapeType variants arranged in grids across 4 slides.

use deckmint::layout::{GridLayoutBuilder, split_v};
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, Presentation, ShapeType};

/// Helper: add a shape with a label below it, given a cell rect.
fn add_labeled_shape(
    slide: &mut deckmint::Slide,
    shape: ShapeType,
    label: &str,
    cell: deckmint::CellRect,
    fill: &str,
) {
    let strips = split_v(&cell, 2, 0.02);
    // Shape in the top portion (inset slightly for padding)
    let shape_area = strips[0].inset(0.06);
    slide.add_shape(
        shape,
        ShapeOptionsBuilder::new()
            .rect(shape_area)
            .fill_color(fill)
            .build(),
    );
    // Label in the bottom strip
    slide.add_text(
        label,
        TextOptionsBuilder::new()
            .rect(strips[1])
            .font_size(7.0)
            .color("#444444")
            .align(AlignH::Center)
            .valign(AlignV::Top)
            .build(),
    );
}

fn main() {
    let mut pres = Presentation::new();

    // Color palettes per slide
    let palette1 = [
        "#4472C4", "#5B9BD5", "#2E75B6", "#1F4E79",
        "#ED7D31", "#F4A460", "#D35400", "#A0522D",
        "#70AD47", "#A9D18E", "#548235", "#2E7D32",
    ];
    let palette2 = [
        "#FFC000", "#FFD54F", "#FF8F00", "#E65100",
        "#9B59B6", "#CE93D8", "#7B1FA2", "#4A148C",
        "#E74C3C", "#EF5350", "#C62828", "#B71C1C",
    ];
    let palette3 = [
        "#1ABC9C", "#2ECC71", "#3498DB", "#9B59B6",
        "#E67E22", "#E74C3C", "#34495E", "#16A085",
        "#27AE60", "#2980B9", "#8E44AD", "#F39C12",
    ];
    let palette4 = [
        "#00BCD4", "#009688", "#4CAF50", "#8BC34A",
        "#CDDC39", "#FFC107", "#FF9800", "#FF5722",
        "#795548", "#9E9E9E", "#607D8B", "#3F51B5",
    ];

    // ══════════════════════════════════════════════════════════
    // Slide 1: Basic shapes
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FAFAFA");

        slide.add_text(
            "Basic Shapes",
            TextOptionsBuilder::new()
                .bounds(0.3, 0.15, 9.4, 0.5)
                .font_size(22.0)
                .bold()
                .color("#1B2A4A")
                .build(),
        );

        let shapes: Vec<(ShapeType, &str)> = vec![
            (ShapeType::Rect, "Rect"),
            (ShapeType::RoundRect, "RoundRect"),
            (ShapeType::Snip1Rect, "Snip1Rect"),
            (ShapeType::Ellipse, "Ellipse"),
            (ShapeType::Triangle, "Triangle"),
            (ShapeType::RtTriangle, "RtTriangle"),
            (ShapeType::Diamond, "Diamond"),
            (ShapeType::Pentagon, "Pentagon"),
            (ShapeType::Hexagon, "Hexagon"),
            (ShapeType::Heptagon, "Heptagon"),
            (ShapeType::Octagon, "Octagon"),
            (ShapeType::Decagon, "Decagon"),
            (ShapeType::Dodecagon, "Dodecagon"),
            (ShapeType::Trapezoid, "Trapezoid"),
            (ShapeType::Parallelogram, "Parallelogram"),
            (ShapeType::Plus, "Plus"),
            (ShapeType::Donut, "Donut"),
            (ShapeType::Can, "Can"),
            (ShapeType::Cube, "Cube"),
            (ShapeType::Bevel, "Bevel"),
            (ShapeType::Frame, "Frame"),
            (ShapeType::Plaque, "Plaque"),
            (ShapeType::Chord, "Chord"),
            (ShapeType::Teardrop, "Teardrop"),
        ];

        let cols = 6;
        let rows = 4;
        let grid = GridLayoutBuilder::grid_n_m(cols, rows, 0.1)
            .origin(0.3, 0.7)
            .container(9.4, 4.75)
            .build();

        for (i, (shape, label)) in shapes.iter().enumerate() {
            let col = i % cols;
            let row = i / cols;
            if row >= rows { break; }
            let cell = grid.cell(col, row);
            let color = palette1[i % palette1.len()];
            add_labeled_shape(slide, shape.clone(), label, cell, color);
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Stars and decorative shapes
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FAFAFA");

        slide.add_text(
            "Stars & Decorative Shapes",
            TextOptionsBuilder::new()
                .bounds(0.3, 0.15, 9.4, 0.5)
                .font_size(22.0)
                .bold()
                .color("#1B2A4A")
                .build(),
        );

        let shapes: Vec<(ShapeType, &str)> = vec![
            (ShapeType::Star4, "Star4"),
            (ShapeType::Star5, "Star5"),
            (ShapeType::Star6, "Star6"),
            (ShapeType::Star7, "Star7"),
            (ShapeType::Star8, "Star8"),
            (ShapeType::Star10, "Star10"),
            (ShapeType::Star12, "Star12"),
            (ShapeType::Star16, "Star16"),
            (ShapeType::Star24, "Star24"),
            (ShapeType::Star32, "Star32"),
            (ShapeType::Heart, "Heart"),
            (ShapeType::Moon, "Moon"),
            (ShapeType::Sun, "Sun"),
            (ShapeType::SmileyFace, "SmileyFace"),
            (ShapeType::LightningBolt, "LightningBolt"),
            (ShapeType::Cloud, "Cloud"),
            (ShapeType::Wave, "Wave"),
            (ShapeType::DoubleWave, "DoubleWave"),
            (ShapeType::Ribbon, "Ribbon"),
            (ShapeType::Ribbon2, "Ribbon2"),
            (ShapeType::NoSmoking, "NoSmoking"),
            (ShapeType::Funnel, "Funnel"),
            (ShapeType::Gear6, "Gear6"),
            (ShapeType::Gear9, "Gear9"),
        ];

        let cols = 6;
        let rows = 4;
        let grid = GridLayoutBuilder::grid_n_m(cols, rows, 0.1)
            .origin(0.3, 0.7)
            .container(9.4, 4.75)
            .build();

        for (i, (shape, label)) in shapes.iter().enumerate() {
            let col = i % cols;
            let row = i / cols;
            if row >= rows { break; }
            let cell = grid.cell(col, row);
            let color = palette2[i % palette2.len()];
            add_labeled_shape(slide, shape.clone(), label, cell, color);
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Arrow shapes
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FAFAFA");

        slide.add_text(
            "Arrow Shapes",
            TextOptionsBuilder::new()
                .bounds(0.3, 0.15, 9.4, 0.5)
                .font_size(22.0)
                .bold()
                .color("#1B2A4A")
                .build(),
        );

        let shapes: Vec<(ShapeType, &str)> = vec![
            (ShapeType::RightArrow, "RightArrow"),
            (ShapeType::LeftArrow, "LeftArrow"),
            (ShapeType::UpArrow, "UpArrow"),
            (ShapeType::DownArrow, "DownArrow"),
            (ShapeType::LeftRightArrow, "LeftRightArrow"),
            (ShapeType::UpDownArrow, "UpDownArrow"),
            (ShapeType::BentArrow, "BentArrow"),
            (ShapeType::BentUpArrow, "BentUpArrow"),
            (ShapeType::UturnArrow, "UturnArrow"),
            (ShapeType::CurvedRightArrow, "CurvedRight"),
            (ShapeType::CurvedLeftArrow, "CurvedLeft"),
            (ShapeType::CurvedUpArrow, "CurvedUp"),
            (ShapeType::CurvedDownArrow, "CurvedDown"),
            (ShapeType::StripedRightArrow, "StripedRight"),
            (ShapeType::NotchedRightArrow, "NotchedRight"),
            (ShapeType::Chevron, "Chevron"),
            (ShapeType::HomePlate, "HomePlate"),
            (ShapeType::QuadArrow, "QuadArrow"),
            (ShapeType::CircularArrow, "CircularArrow"),
            (ShapeType::SwooshArrow, "SwooshArrow"),
            (ShapeType::LeftUpArrow, "LeftUpArrow"),
            (ShapeType::LeftRightUpArrow, "LRUpArrow"),
            (ShapeType::BlockArc, "BlockArc"),
            (ShapeType::Arc, "Arc"),
        ];

        let cols = 6;
        let rows = 4;
        let grid = GridLayoutBuilder::grid_n_m(cols, rows, 0.1)
            .origin(0.3, 0.7)
            .container(9.4, 4.75)
            .build();

        for (i, (shape, label)) in shapes.iter().enumerate() {
            let col = i % cols;
            let row = i / cols;
            if row >= rows { break; }
            let cell = grid.cell(col, row);
            let color = palette3[i % palette3.len()];
            add_labeled_shape(slide, shape.clone(), label, cell, color);
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Flowchart shapes
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FAFAFA");

        slide.add_text(
            "Flowchart Shapes",
            TextOptionsBuilder::new()
                .bounds(0.3, 0.15, 9.4, 0.5)
                .font_size(22.0)
                .bold()
                .color("#1B2A4A")
                .build(),
        );

        let shapes: Vec<(ShapeType, &str)> = vec![
            (ShapeType::FlowChartProcess, "Process"),
            (ShapeType::FlowChartAlternateProcess, "AltProcess"),
            (ShapeType::FlowChartDecision, "Decision"),
            (ShapeType::FlowChartInputOutput, "Input/Output"),
            (ShapeType::FlowChartPredefinedProcess, "Predefined"),
            (ShapeType::FlowChartDocument, "Document"),
            (ShapeType::FlowChartMultidocument, "MultiDoc"),
            (ShapeType::FlowChartTerminator, "Terminator"),
            (ShapeType::FlowChartPreparation, "Preparation"),
            (ShapeType::FlowChartManualInput, "ManualInput"),
            (ShapeType::FlowChartManualOperation, "ManualOp"),
            (ShapeType::FlowChartConnector, "Connector"),
            (ShapeType::FlowChartOffpageConnector, "OffpageConn"),
            (ShapeType::FlowChartDelay, "Delay"),
            (ShapeType::FlowChartDisplay, "Display"),
            (ShapeType::FlowChartSort, "Sort"),
            (ShapeType::FlowChartMerge, "Merge"),
            (ShapeType::FlowChartExtract, "Extract"),
            (ShapeType::FlowChartCollate, "Collate"),
            (ShapeType::FlowChartOr, "Or"),
            (ShapeType::FlowChartInternalStorage, "IntStorage"),
            (ShapeType::FlowChartMagneticDisk, "MagDisk"),
            (ShapeType::FlowChartPunchedCard, "PunchedCard"),
            (ShapeType::FlowChartPunchedTape, "PunchedTape"),
        ];

        let cols = 6;
        let rows = 4;
        let grid = GridLayoutBuilder::grid_n_m(cols, rows, 0.1)
            .origin(0.3, 0.7)
            .container(9.4, 4.75)
            .build();

        for (i, (shape, label)) in shapes.iter().enumerate() {
            let col = i % cols;
            let row = i / cols;
            if row >= rows { break; }
            let cell = grid.cell(col, row);
            let color = palette4[i % palette4.len()];
            add_labeled_shape(slide, shape.clone(), label, cell, color);
        }
    }

    pres.write_to_file("21_shapes_showcase.pptx").unwrap();
    println!("Wrote 21_shapes_showcase.pptx");
}
