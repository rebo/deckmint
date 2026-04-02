//! Grid layout helper for arranging objects in rows and columns.

use deckmint::layout::GridLayoutBuilder;
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();

    slide.add_text(
        "Grid Layout Demo",
        TextOptionsBuilder::new()
            .pos(0.5, 0.2).size(9.0, 0.7)
            .font_size(24.0).bold()
            .build(),
    );

    // 2 rows × 3 columns with gutters
    let grid = GridLayoutBuilder::grid_n_m(3, 2, 0.2)
        .origin(0.5, 1.2)
        .container(9.0, 4.2)
        .build();

    let colors = ["4472C4", "ED7D31", "70AD47", "FFC000", "5B9BD5", "A9D18E"];
    let labels = ["Cell 1", "Cell 2", "Cell 3", "Cell 4", "Cell 5", "Cell 6"];

    for row in 0..2 {
        for col in 0..3 {
            let idx = row * 3 + col;
            let cell = grid.cell(col, row);

            slide.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .pos(cell.x, cell.y).size(cell.w, cell.h)
                    .fill_color(colors[idx])
                    .rect_radius(0.1)
                    .build(),
            );
            slide.add_text(
                labels[idx],
                TextOptionsBuilder::new()
                    .pos(cell.x, cell.y).size(cell.w, cell.h)
                    .font_size(18.0).bold()
                    .color("FFFFFF")
                    .align(AlignH::Center).valign(AlignV::Middle)
                    .build(),
            );
        }
    }

    pres.write_to_file("09_grid_layout.pptx").unwrap();
    println!("Wrote 09_grid_layout.pptx");
}
