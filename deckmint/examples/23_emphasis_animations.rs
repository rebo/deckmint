//! Emphasis animation effects: spin, pulse, grow/shrink, teeter, shimmer, blink, and more.

use deckmint::layout::GridLayoutBuilder;
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, AnimationEffect, Presentation, ShapeType};

/// Helper: add a labeled shape with an emphasis animation.
fn add_anim_shape(
    slide: &mut deckmint::Slide,
    cell: deckmint::CellRect,
    fill: &str,
    label: &str,
    anim: AnimationEffect,
) {
    let (top, bottom) = cell.halves_v(0.05);
    let shape_area = top.inset(0.08);

    slide.add_shape(
        ShapeType::RoundRect,
        ShapeOptionsBuilder::new()
            .rect(shape_area)
            .fill_color(fill)
            .rect_radius(0.08)
            .animation(anim)
            .build(),
    );

    // White label centered inside the shape
    slide.add_text(
        label,
        TextOptionsBuilder::new()
            .rect(shape_area)
            .font_size(10.0)
            .bold()
            .color("#FFFFFF")
            .align(AlignH::Center)
            .valign(AlignV::Middle)
            .build(),
    );

    // Description label below the shape
    slide.add_text(
        label,
        TextOptionsBuilder::new()
            .rect(bottom)
            .font_size(8.0)
            .color("#555555")
            .align(AlignH::Center)
            .valign(AlignV::Top)
            .build(),
    );
}

fn main() {
    let mut pres = Presentation::new();

    // ══════════════════════════════════════════════════════════
    // Slide 1: Basic emphasis animations
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#F8F9FA");

        slide.add_text(
            "Basic Emphasis Animations",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        slide.add_text(
            "Click each shape to trigger its emphasis animation",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.6, 9.0, 0.3)
                .font_size(11.0)
                .color("#888888")
                .align(AlignH::Center)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(4, 2, 0.15)
            .origin(0.5, 1.0)
            .container(9.0, 4.3)
            .build();

        let anims: Vec<(&str, &str, AnimationEffect)> = vec![
            ("Spin 360",     "#4472C4", AnimationEffect::spin(360.0)),
            ("Pulse",        "#ED7D31", AnimationEffect::pulse()),
            ("Grow 150%",    "#70AD47", AnimationEffect::grow_shrink(1.5)),
            ("Shrink 50%",   "#FFC000", AnimationEffect::grow_shrink(0.5)),
            ("Teeter",       "#5B9BD5", AnimationEffect::teeter()),
            ("Shimmer",      "#9B59B6", AnimationEffect::shimmer()),
            ("Blink",        "#E74C3C", AnimationEffect::blink()),
            ("Spin 720",     "#1ABC9C", AnimationEffect::spin(720.0)),
        ];

        for (i, (label, color, anim)) in anims.into_iter().enumerate() {
            let col = i % 4;
            let row = i / 4;
            add_anim_shape(slide, grid.cell(col, row), color, label, anim);
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Color-based emphasis animations
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#F8F9FA");

        slide.add_text(
            "Color-Based Emphasis",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        slide.add_text(
            "These animations change the color properties of shapes",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.6, 9.0, 0.3)
                .font_size(11.0)
                .color("#888888")
                .align(AlignH::Center)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(3, 2, 0.2)
            .origin(0.8, 1.1)
            .container(8.4, 4.2)
            .build();

        // Fill color change
        {
            let cell = grid.cell(0, 0);
            let (top, bottom) = cell.halves_v(0.05);
            let shape_area = top.inset(0.08);
            slide.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(shape_area)
                    .fill_color("#4472C4")
                    .rect_radius(0.08)
                    .animation(AnimationEffect::fill_color("#E74C3C"))
                    .build(),
            );
            slide.add_text("Fill Color", TextOptionsBuilder::new()
                .rect(shape_area).font_size(11.0).bold().color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle).build());
            slide.add_text("Blue -> Red", TextOptionsBuilder::new()
                .rect(bottom).font_size(8.0).color("#555555")
                .align(AlignH::Center).valign(AlignV::Top).build());
        }

        // Font color change
        {
            let cell = grid.cell(1, 0);
            let (top, bottom) = cell.halves_v(0.05);
            let shape_area = top.inset(0.08);
            slide.add_text(
                "HELLO",
                TextOptionsBuilder::new()
                    .rect(shape_area)
                    .font_size(28.0)
                    .bold()
                    .color("#1B2A4A")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .animation(AnimationEffect::font_color("#E74C3C"))
                    .build(),
            );
            slide.add_text("Font Color", TextOptionsBuilder::new()
                .rect(bottom).font_size(8.0).color("#555555")
                .align(AlignH::Center).valign(AlignV::Top).build());
        }

        // Line color change
        {
            let cell = grid.cell(2, 0);
            let (top, bottom) = cell.halves_v(0.05);
            let shape_area = top.inset(0.08);
            slide.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(shape_area)
                    .fill_color("#F0F0F0")
                    .line_color("#4472C4")
                    .line_width(3.0)
                    .rect_radius(0.08)
                    .animation(AnimationEffect::line_color("#E74C3C"))
                    .build(),
            );
            slide.add_text("Line Color", TextOptionsBuilder::new()
                .rect(shape_area).font_size(11.0).bold().color("#333333")
                .align(AlignH::Center).valign(AlignV::Middle).build());
            slide.add_text("Blue -> Red border", TextOptionsBuilder::new()
                .rect(bottom).font_size(8.0).color("#555555")
                .align(AlignH::Center).valign(AlignV::Top).build());
        }

        // Bold flash
        {
            let cell = grid.cell(0, 1);
            let (top, bottom) = cell.halves_v(0.05);
            let shape_area = top.inset(0.08);
            slide.add_text(
                "Bold Flash",
                TextOptionsBuilder::new()
                    .rect(shape_area)
                    .font_size(20.0)
                    .color("#2C3E50")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .animation(AnimationEffect::bold_flash())
                    .build(),
            );
            slide.add_text("Bold Flash", TextOptionsBuilder::new()
                .rect(bottom).font_size(8.0).color("#555555")
                .align(AlignH::Center).valign(AlignV::Top).build());
        }

        // Bold reveal
        {
            let cell = grid.cell(1, 1);
            let (top, bottom) = cell.halves_v(0.05);
            let shape_area = top.inset(0.08);
            slide.add_text(
                "Bold Reveal",
                TextOptionsBuilder::new()
                    .rect(shape_area)
                    .font_size(20.0)
                    .color("#2C3E50")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .animation(AnimationEffect::bold_reveal())
                    .build(),
            );
            slide.add_text("Bold Reveal", TextOptionsBuilder::new()
                .rect(bottom).font_size(8.0).color("#555555")
                .align(AlignH::Center).valign(AlignV::Top).build());
        }

        // Wave
        {
            let cell = grid.cell(2, 1);
            let (top, bottom) = cell.halves_v(0.05);
            let shape_area = top.inset(0.08);
            slide.add_text(
                "Wave Effect",
                TextOptionsBuilder::new()
                    .rect(shape_area)
                    .font_size(20.0)
                    .color("#2C3E50")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .animation(AnimationEffect::wave())
                    .build(),
            );
            slide.add_text("Wave", TextOptionsBuilder::new()
                .rect(bottom).font_size(8.0).color("#555555")
                .align(AlignH::Center).valign(AlignV::Top).build());
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Combining emphasis with shape variety
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#1A1A2E");

        slide.add_text(
            "Emphasis on Various Shapes",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color("#E0E0E0")
                .align(AlignH::Center)
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(5, 2, 0.15)
            .origin(0.4, 0.8)
            .container(9.2, 4.5)
            .build();

        let shape_anims: Vec<(ShapeType, &str, &str, AnimationEffect)> = vec![
            (ShapeType::Ellipse,    "#E74C3C", "Spin",      AnimationEffect::spin(360.0)),
            (ShapeType::Diamond,    "#3498DB", "Pulse",     AnimationEffect::pulse()),
            (ShapeType::Star5,      "#F1C40F", "Grow",      AnimationEffect::grow_shrink(1.3)),
            (ShapeType::Hexagon,    "#2ECC71", "Teeter",    AnimationEffect::teeter()),
            (ShapeType::Heart,      "#E91E63", "Shimmer",   AnimationEffect::shimmer()),
            (ShapeType::Pentagon,   "#9B59B6", "Blink",     AnimationEffect::blink()),
            (ShapeType::Octagon,    "#1ABC9C", "Spin 180",  AnimationEffect::spin(180.0)),
            (ShapeType::Star8,      "#E67E22", "Shrink",    AnimationEffect::grow_shrink(0.6)),
            (ShapeType::Moon,       "#5DADE2", "Fill Clr",  AnimationEffect::fill_color("#FF6B6B")),
            (ShapeType::Cloud,      "#AED6F1", "Wave",      AnimationEffect::wave()),
        ];

        for (i, (shape, color, label, anim)) in shape_anims.into_iter().enumerate() {
            let col = i % 5;
            let row = i / 5;
            let cell = grid.cell(col, row);
            let (top, bottom) = cell.halves_v(0.05);
            let shape_area = top.inset(0.1);

            slide.add_shape(
                shape,
                ShapeOptionsBuilder::new()
                    .rect(shape_area)
                    .fill_color(color)
                    .animation(anim)
                    .build(),
            );

            slide.add_text(
                label,
                TextOptionsBuilder::new()
                    .rect(bottom)
                    .font_size(9.0)
                    .color("#AAAAAA")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }
    }

    pres.write_to_file("23_emphasis_animations.pptx").unwrap();
    println!("Wrote 23_emphasis_animations.pptx");
}
