//! Theme color system: scheme colors, tint/shade modifiers, and themed text.

use deckmint::layout::{GridLayoutBuilder, GridTrack};
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::types::ThemeColorMod;
use deckmint::{AlignH, AlignV, Color, Presentation, SchemeColor, ShapeType};

fn main() {
    let mut pres = Presentation::new();

    // ══════════════════════════════════════════════════════════
    // Slide 1: The 6 accent scheme colors as large swatches
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FFFFFF");

        slide.add_text(
            "Theme Accent Colors",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(26.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        slide.add_text(
            "These colors adapt automatically when you change the presentation theme.",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.75, 9.0, 0.35)
                .font_size(12.0)
                .color("#666666")
                .align(AlignH::Center)
                .build(),
        );

        let accents: Vec<(SchemeColor, &str)> = vec![
            (SchemeColor::Accent1, "Accent 1"),
            (SchemeColor::Accent2, "Accent 2"),
            (SchemeColor::Accent3, "Accent 3"),
            (SchemeColor::Accent4, "Accent 4"),
            (SchemeColor::Accent5, "Accent 5"),
            (SchemeColor::Accent6, "Accent 6"),
        ];

        let grid = GridLayoutBuilder::grid_n_m(6, 1, 0.2)
            .origin(0.5, 1.3)
            .container(9.0, 3.0)
            .build();

        for (i, (scheme, label)) in accents.iter().enumerate() {
            let cell = grid.cell(i, 0);

            // Large swatch rectangle
            slide.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell.inset_xy(0.0, 0.0))
                    .fill_color_value(Color::Theme(scheme.clone()))
                    .rect_radius(0.1)
                    .build(),
            );

            // Label inside the swatch
            slide.add_text(
                *label,
                TextOptionsBuilder::new()
                    .rect(cell)
                    .font_size(14.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Also show Text and Background scheme colors
        let bottom_grid = GridLayoutBuilder::grid_n_m(4, 1, 0.2)
            .origin(1.5, 4.5)
            .container(7.0, 0.8)
            .build();

        let bg_colors: Vec<(SchemeColor, &str, &str)> = vec![
            (SchemeColor::Text1, "Text 1", "#FFFFFF"),
            (SchemeColor::Text2, "Text 2", "#FFFFFF"),
            (SchemeColor::Background1, "Background 1", "#333333"),
            (SchemeColor::Background2, "Background 2", "#333333"),
        ];

        for (i, (scheme, label, txt_color)) in bg_colors.iter().enumerate() {
            let cell = bottom_grid.cell(i, 0);
            slide.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell)
                    .fill_color_value(Color::Theme(scheme.clone()))
                    .rect_radius(0.06)
                    .line_color("#CCCCCC")
                    .line_width(0.5)
                    .build(),
            );
            slide.add_text(
                *label,
                TextOptionsBuilder::new()
                    .rect(cell)
                    .font_size(10.0)
                    .bold()
                    .color(*txt_color)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Theme color modifiers — tint, shade, luminance
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FFFFFF");

        slide.add_text(
            "Theme Color Modifiers",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        // Show Accent1 with different tint and shade levels
        let accents_to_show = [
            (SchemeColor::Accent1, "Accent 1"),
            (SchemeColor::Accent2, "Accent 2"),
            (SchemeColor::Accent5, "Accent 5"),
        ];

        let row_grid = GridLayoutBuilder::new()
            .cols(vec![GridTrack::Inches(1.3), GridTrack::Fr(1.0)])
            .rows(vec![GridTrack::Fr(1.0); 3])
            .gap(0.12)
            .origin(0.4, 0.75)
            .container(9.2, 4.6)
            .build();

        for (row_idx, (scheme, label)) in accents_to_show.iter().enumerate() {
            // Row label
            let label_cell = row_grid.cell(0, row_idx);
            slide.add_text(
                *label,
                TextOptionsBuilder::new()
                    .rect(label_cell)
                    .font_size(12.0)
                    .bold()
                    .color("#333333")
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Swatches area
            let swatch_area = row_grid.cell(1, row_idx);
            let swatch_grid = GridLayoutBuilder::within(swatch_area)
                .cols(vec![GridTrack::Fr(1.0); 7])
                .rows(vec![GridTrack::Fr(1.0), GridTrack::Inches(0.3)])
                .col_gap(0.08)
                .build();

            // Tint 80%, Tint 60%, Tint 40%, Base, Shade 25%, Shade 50%, Shade 75%
            let modifiers: Vec<(Option<ThemeColorMod>, &str)> = vec![
                (Some(ThemeColorMod::tint(20000)), "Tint 80%"),
                (Some(ThemeColorMod::tint(40000)), "Tint 60%"),
                (Some(ThemeColorMod::tint(60000)), "Tint 40%"),
                (None, "Base"),
                (Some(ThemeColorMod::shade(75000)), "Shade 25%"),
                (Some(ThemeColorMod::shade(50000)), "Shade 50%"),
                (Some(ThemeColorMod::shade(25000)), "Shade 75%"),
            ];

            for (col, (modifier, mod_label)) in modifiers.iter().enumerate() {
                let swatch_cell = swatch_grid.cell(col, 0);
                let color = match modifier {
                    Some(m) => Color::ThemedWith(scheme.clone(), m.clone()),
                    None => Color::Theme(scheme.clone()),
                };

                slide.add_shape(
                    ShapeType::RoundRect,
                    ShapeOptionsBuilder::new()
                        .rect(swatch_cell.inset(0.02))
                        .fill_color_value(color)
                        .rect_radius(0.06)
                        .build(),
                );

                let label_cell = swatch_grid.cell(col, 1);
                slide.add_text(
                    *mod_label,
                    TextOptionsBuilder::new()
                        .rect(label_cell)
                        .font_size(7.0)
                        .color("#666666")
                        .align(AlignH::Center)
                        .valign(AlignV::Top)
                        .build(),
                );
            }
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Luminance modifiers and practical color palettes
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#FFFFFF");

        slide.add_text(
            "Luminance Modifiers",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(24.0)
                .bold()
                .color("#1B2A4A")
                .align(AlignH::Center)
                .build(),
        );

        slide.add_text(
            "Luminance modifiers adjust brightness via lumMod and lumOff OOXML attributes",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.6, 9.0, 0.35)
                .font_size(11.0)
                .color("#888888")
                .align(AlignH::Center)
                .build(),
        );

        // Show Accent1 through Accent6 with luminance variations
        let all_accents = [
            SchemeColor::Accent1,
            SchemeColor::Accent2,
            SchemeColor::Accent3,
            SchemeColor::Accent4,
            SchemeColor::Accent5,
            SchemeColor::Accent6,
        ];

        let lum_grid = GridLayoutBuilder::grid_n_m(6, 5, 0.08)
            .origin(0.5, 1.1)
            .container(9.0, 3.4)
            .build();

        // Luminance variations: lighter to darker
        let lum_variants: Vec<(ThemeColorMod, &str)> = vec![
            (ThemeColorMod::lum(20000, 80000), "Lighter 80%"),
            (ThemeColorMod::lum(40000, 60000), "Lighter 60%"),
            (ThemeColorMod::lum(60000, 40000), "Lighter 40%"),
            (ThemeColorMod::lum(75000, 0), "Darker 25%"),
            (ThemeColorMod::lum(50000, 0), "Darker 50%"),
        ];

        // Column headers
        for (col, accent) in all_accents.iter().enumerate() {
            // Full-color header
            let cell = lum_grid.cell(col, 0);
            slide.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell.inset(0.02))
                    .fill_color_value(Color::Theme(accent.clone()))
                    .rect_radius(0.04)
                    .build(),
            );
            slide.add_text(
                &format!("Accent {}", col + 1),
                TextOptionsBuilder::new()
                    .rect(cell)
                    .font_size(9.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Luminance rows (skip the first grid row, it's the header)
        // Rearrange: use row 1-4 for lum variants (only first 4)
        for (row_offset, (lum, _label)) in lum_variants.iter().take(4).enumerate() {
            let row = row_offset + 1;
            for (col, accent) in all_accents.iter().enumerate() {
                let cell = lum_grid.cell(col, row);
                slide.add_shape(
                    ShapeType::RoundRect,
                    ShapeOptionsBuilder::new()
                        .rect(cell.inset(0.02))
                        .fill_color_value(Color::ThemedWith(accent.clone(), lum.clone()))
                        .rect_radius(0.04)
                        .build(),
                );
            }
        }

        // Row labels on the left side
        let row_labels = ["Base", "Lighter 80%", "Lighter 60%", "Lighter 40%", "Darker 25%"];
        for (i, label) in row_labels.iter().enumerate() {
            let y = 1.1 + (i as f64) * (3.4 / 5.0);
            slide.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(0.0, y, 0.5, 3.4 / 5.0)
                    .font_size(6.0)
                    .color("#999999")
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Bottom: practical example using themed colors for a card layout
        slide.add_text(
            "Theme colors make your presentation portable across themes",
            TextOptionsBuilder::new()
                .bounds(0.5, 4.7, 9.0, 0.5)
                .font_size(12.0)
                .color("#555555")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    pres.write_to_file("22_theme_colors.pptx").unwrap();
    println!("Wrote 22_theme_colors.pptx");
}
