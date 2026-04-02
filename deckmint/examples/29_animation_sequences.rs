//! Complex animation sequences: AfterPrevious triggers with delays,
//! WithPrevious (simultaneous) animations, multiple click groups,
//! and mixed entrance/exit animations.

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{AlignH, AlignV, AnimationEffect, Direction, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Animation Sequences".to_string();

    // ══════════════════════════════════════════════════════════
    // Slide 1: Objects appearing one by one (AfterPrevious)
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#1B2A4A");

        s.add_text(
            "Sequential Animations",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(28.0)
                .bold()
                .color("#FFFFFF")
                .build(),
        );

        s.add_text(
            "Click once, then watch objects appear one after another",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.75, 9.0, 0.4)
                .font_size(13.0)
                .color("#8FAADC")
                .build(),
        );

        let steps = [
            ("Step 1", "Discover", "#4472C4"),
            ("Step 2", "Design", "#ED7D31"),
            ("Step 3", "Develop", "#70AD47"),
            ("Step 4", "Deploy", "#FFC000"),
            ("Step 5", "Deliver", "#5B9BD5"),
        ];

        let card_w = 1.6;
        let card_h = 2.5;
        let gap = 0.2;
        let total_w = steps.len() as f64 * card_w + (steps.len() as f64 - 1.0) * gap;
        let x_start = (10.0 - total_w) / 2.0;
        let y_pos = 1.5;

        for (i, (step, label, color)) in steps.iter().enumerate() {
            let x = x_start + i as f64 * (card_w + gap);

            // Each card flies in from the bottom, one after another
            let anim = if i == 0 {
                // First one triggers on click
                AnimationEffect::fly_in(Direction::Down)
            } else {
                // Rest follow automatically with staggered delay
                AnimationEffect::fly_in(Direction::Down)
                    .after_previous()
                    .delay(300)
            };

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(x, y_pos, card_w, card_h)
                    .fill_color(*color)
                    .rect_radius(0.1)
                    .animation(anim)
                    .build(),
            );

            // Step number circle
            let cx = x + (card_w - 0.5) / 2.0;
            let anim_circle = if i == 0 {
                AnimationEffect::zoom_in()
            } else {
                AnimationEffect::zoom_in().after_previous().delay(0)
            };

            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(cx, y_pos + 0.3, 0.5, 0.5)
                    .fill_color("#FFFFFF")
                    .animation(anim_circle)
                    .build(),
            );
            s.add_text(
                &format!("{}", i + 1),
                TextOptionsBuilder::new()
                    .bounds(cx, y_pos + 0.3, 0.5, 0.5)
                    .font_size(18.0)
                    .bold()
                    .color(*color)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Step label
            s.add_text(
                *step,
                TextOptionsBuilder::new()
                    .bounds(x, y_pos + 1.0, card_w, 0.4)
                    .font_size(11.0)
                    .color("#E0E8F0")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(x, y_pos + 1.4, card_w, 0.6)
                    .font_size(18.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }

        // Completion message appears last
        s.add_text(
            "Process Complete!",
            TextOptionsBuilder::new()
                .bounds(2.0, 4.3, 6.0, 0.6)
                .font_size(22.0)
                .bold()
                .color("#70AD47")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .animation(AnimationEffect::fade_in().after_previous().delay(500))
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Objects appearing together (WithPrevious)
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#0D1B2A");

        s.add_text(
            "Simultaneous Animations",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(28.0)
                .bold()
                .color("#FFFFFF")
                .build(),
        );

        s.add_text(
            "Click once to see all shapes animate at the same time",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.75, 9.0, 0.4)
                .font_size(13.0)
                .color("#90E0EF")
                .build(),
        );

        // Group 1: Four shapes fly in from different directions simultaneously
        let shapes_g1 = [
            ("Top", Direction::Up, 1.5, 1.5, "#4472C4"),
            ("Left", Direction::Left, 1.5, 1.5, "#ED7D31"),
            ("Right", Direction::Right, 1.5, 1.5, "#70AD47"),
            ("Bottom", Direction::Down, 1.5, 1.5, "#FFC000"),
        ];

        let positions = [
            (3.5, 1.4),   // top-center
            (1.0, 2.6),   // left
            (6.0, 2.6),   // right
            (3.5, 3.8),   // bottom-center
        ];

        for (i, ((label, dir, w, h, color), (x, y))) in
            shapes_g1.iter().zip(positions.iter()).enumerate()
        {
            let anim = if i == 0 {
                // First one triggers on click
                AnimationEffect::fly_in(dir.clone())
            } else {
                // Others play simultaneously
                AnimationEffect::fly_in(dir.clone()).with_previous()
            };

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(*x, *y, *w, *h)
                    .fill_color(*color)
                    .rect_radius(0.1)
                    .animation(anim)
                    .build(),
            );

            s.add_text(
                &format!("Fly from\n{}", label),
                TextOptionsBuilder::new()
                    .bounds(*x, *y, *w, *h)
                    .font_size(14.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Center circle pulses after all fly-ins complete
        s.add_shape(
            ShapeType::Ellipse,
            ShapeOptionsBuilder::new()
                .bounds(4.0, 2.75, 2.0, 2.0)
                .fill_color("#9B59B6")
                .animation(AnimationEffect::zoom_in().after_previous().delay(200))
                .build(),
        );
        s.add_text(
            "CENTER",
            TextOptionsBuilder::new()
                .bounds(4.0, 2.75, 2.0, 2.0)
                .font_size(16.0)
                .bold()
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Multiple click groups
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#F5F6FA");

        s.add_text(
            "Multiple Click Groups",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(28.0)
                .bold()
                .color("#1B2A4A")
                .build(),
        );

        s.add_text(
            "Each click reveals a different group of shapes",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.6, 9.0, 0.35)
                .font_size(13.0)
                .color("#666666")
                .build(),
        );

        // Group labels
        let groups = [
            ("Click 1: Revenue", "#4472C4", 1u32),
            ("Click 2: Growth", "#70AD47", 2),
            ("Click 3: Forecast", "#ED7D31", 3),
        ];

        let col_w = 2.8;
        let col_gap = 0.25;
        let x_start = 0.55;

        for (gi, (group_title, color, group_num)) in groups.iter().enumerate() {
            let gx = x_start + gi as f64 * (col_w + col_gap);

            // Group header
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(gx, 1.1, col_w, 0.5)
                    .fill_color(*color)
                    .rect_radius(0.06)
                    .animation(AnimationEffect::fade_in().with_group(*group_num))
                    .build(),
            );
            s.add_text(
                *group_title,
                TextOptionsBuilder::new()
                    .bounds(gx, 1.1, col_w, 0.5)
                    .font_size(14.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Three data bars in each group
            let bar_data = match gi {
                0 => vec![("Q1", 0.6), ("Q2", 0.75), ("Q3", 0.9)],
                1 => vec![("North", 0.8), ("South", 0.5), ("East", 0.65)],
                _ => vec![("2025", 0.7), ("2026", 0.85), ("2027", 1.0)],
            };

            for (bi, (bar_label, bar_pct)) in bar_data.iter().enumerate() {
                let by = 1.85 + bi as f64 * 0.95;
                let bar_w = (col_w - 0.4) * bar_pct;

                // Label
                s.add_text(
                    *bar_label,
                    TextOptionsBuilder::new()
                        .bounds(gx + 0.1, by, col_w - 0.2, 0.3)
                        .font_size(11.0)
                        .bold()
                        .color("#555555")
                        .build(),
                );

                // Bar background
                s.add_shape(
                    ShapeType::RoundRect,
                    ShapeOptionsBuilder::new()
                        .bounds(gx + 0.1, by + 0.3, col_w - 0.2, 0.35)
                        .fill_color("#E8ECF1")
                        .rect_radius(0.04)
                        .animation(
                            AnimationEffect::fade_in()
                                .with_group(*group_num)
                                .with_previous(),
                        )
                        .build(),
                );

                // Bar fill (animates with group)
                s.add_shape(
                    ShapeType::RoundRect,
                    ShapeOptionsBuilder::new()
                        .bounds(gx + 0.1, by + 0.3, bar_w, 0.35)
                        .fill_color(*color)
                        .rect_radius(0.04)
                        .animation(
                            AnimationEffect::wipe_in(Direction::Left)
                                .with_group(*group_num)
                                .with_previous(),
                        )
                        .build(),
                );

                // Percentage label
                s.add_text(
                    &format!("{:.0}%", bar_pct * 100.0),
                    TextOptionsBuilder::new()
                        .bounds(gx + 0.1, by + 0.3, bar_w, 0.35)
                        .font_size(10.0)
                        .bold()
                        .color("#FFFFFF")
                        .align(AlignH::Right)
                        .valign(AlignV::Middle)
                        .build(),
                );
            }
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Mixed entrance and exit animations
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#0B0C10");

        s.add_text(
            "Entrance & Exit Animations",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(28.0)
                .bold()
                .color("#66FCF1")
                .build(),
        );

        s.add_text(
            "Click to cycle through: appear, emphasize, then disappear",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.6, 9.0, 0.35)
                .font_size(13.0)
                .color("#C5C6C7")
                .build(),
        );

        // Row 1: Entrance animations
        s.add_text(
            "ENTRANCE",
            TextOptionsBuilder::new()
                .bounds(0.5, 1.2, 1.5, 0.3)
                .font_size(10.0)
                .bold()
                .color("#66FCF1")
                .build(),
        );

        let entrances: Vec<(&str, AnimationEffect, &str)> = vec![
            ("Fade In", AnimationEffect::fade_in(), "#4472C4"),
            ("Fly In", AnimationEffect::fly_in(Direction::Left), "#ED7D31"),
            ("Zoom In", AnimationEffect::zoom_in(), "#70AD47"),
            ("Bounce In", AnimationEffect::bounce_in(), "#FFC000"),
        ];

        let item_w = 2.0;
        let item_h = 1.1;
        let item_gap = 0.25;
        let ix_start = 0.5;

        for (i, (label, anim, color)) in entrances.iter().enumerate() {
            let x = ix_start + i as f64 * (item_w + item_gap);

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(x, 1.6, item_w, item_h)
                    .fill_color(*color)
                    .rect_radius(0.08)
                    .animation(anim.clone())
                    .build(),
            );
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(x, 1.6, item_w, item_h)
                    .font_size(14.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Row 2: Exit animations
        s.add_text(
            "EXIT",
            TextOptionsBuilder::new()
                .bounds(0.5, 3.0, 1.5, 0.3)
                .font_size(10.0)
                .bold()
                .color("#E74C3C")
                .build(),
        );

        let exits: Vec<(&str, AnimationEffect, &str)> = vec![
            ("Fade Out", AnimationEffect::fade_out(), "#5B9BD5"),
            ("Fly Out", AnimationEffect::fly_out(Direction::Right), "#9B59B6"),
            ("Zoom Out", AnimationEffect::zoom_out(), "#E74C3C"),
            ("Disappear", AnimationEffect::disappear(), "#1ABC9C"),
        ];

        for (i, (label, anim, color)) in exits.iter().enumerate() {
            let x = ix_start + i as f64 * (item_w + item_gap);

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(x, 3.4, item_w, item_h)
                    .fill_color(*color)
                    .rect_radius(0.08)
                    .animation(anim.clone())
                    .build(),
            );
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(x, 3.4, item_w, item_h)
                    .font_size(14.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Row 3: Emphasis label
        s.add_text(
            "EMPHASIS",
            TextOptionsBuilder::new()
                .bounds(0.5, 4.7, 1.5, 0.3)
                .font_size(10.0)
                .bold()
                .color("#FFC000")
                .build(),
        );

        let emphases: Vec<(&str, AnimationEffect, &str)> = vec![
            ("Spin", AnimationEffect::spin(360.0), "#4472C4"),
            ("Pulse", AnimationEffect::pulse(), "#ED7D31"),
            ("Grow", AnimationEffect::grow_shrink(1.5), "#70AD47"),
        ];

        for (i, (label, anim, color)) in emphases.iter().enumerate() {
            let x = ix_start + i as f64 * (item_w + item_gap);

            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(x + 0.3, 4.95, item_w - 0.6, 0.55)
                    .fill_color(*color)
                    .animation(anim.clone())
                    .build(),
            );
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(x + 0.3, 4.95, item_w - 0.6, 0.55)
                    .font_size(12.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }
    }

    pres.write_to_file("29_animation_sequences.pptx").unwrap();
    println!("Wrote 29_animation_sequences.pptx");
}
