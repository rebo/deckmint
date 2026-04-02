//! Generates `features.pptx` — one slide per feature group, opening in
//! PowerPoint/LibreOffice should show every new capability at a glance.
//!
//! Run:  cargo run --example features

use deckmint::layout::{center_in, GridLayoutBuilder, GridTrack};
use deckmint::{GradientFill, GradientStop};
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::table::{TableCell, TableOptionsBuilder};
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{BarDir, BarGrouping, ChartOptionsBuilder, ChartSeries, LegendPos};
use deckmint::TextRunBuilder;
use deckmint::types::{AnimationEffect, CheckerboardDir, Direction, HyperlinkProps, ShapeVariant, SlideNumberProps, SplitOrientation, StripDir};
use deckmint::{
    AlignH, AlignV, ChartType, GlowProps, Presentation, SchemeColor, ShapeType, SlideMasterDef, TabStop,
    TextOutlineProps,
};
use deckmint::{
    AnimationTrigger, ConnectorOptions, ConnectorType, FieldType, HyperlinkAction,
    PatternFill, PatternType, ThemeColorMod, TransitionDir, TransitionProps,
};

/// Minimal 1×1 red-pixel PNG (used for background-image demo)
const RED_1PX_PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
    0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
    0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8,
    0xCF, 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33, 0x00, 0x00, 0x00,
    0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn heading(slide: &mut deckmint::Slide, text: &str) {
    slide.add_text(
        text,
        TextOptionsBuilder::new()
            .x(0.3).y(0.1).w(9.4).h(0.7)
            .font_size(24.0).bold().color("#2E4057")
            .build(),
    );
    slide.add_shape(
        ShapeType::Rect,
        ShapeOptionsBuilder::new()
            .x(0.3).y(0.75).w(9.4).h(0.04)
            .fill_color("#2E4057")
            .build(),
    );
}

fn main() {
    let mut pres = Presentation::new();
    pres.title = "deckmint Feature Demo".to_string();
    pres.author = "deckmint".to_string();

    // ── Slide master: branded background stripe ─────────────
    {
        let mut master = SlideMasterDef::new("Feature Demo Master");
        master.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
            .bounds(0.0, 7.3, 13.33, 0.2)
            .fill_color("#2E4057")
            .build());
        pres.define_master(master);
    }

    // ── Custom layout: widescreen 16×9 with explicit dims ───
    pres.define_layout("WIDE_16x9", Some(13.33), Some(7.5));

    // ══════════════════════════════════════════════════════════
    // Slide 1: Text Highlight + Underline Styles
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Text: Highlight & Underline Styles");

        // Highlight
        let runs = vec![
            TextRunBuilder::new("Normal · ").font_size(20.0).build(),
            TextRunBuilder::new("Yellow highlight").font_size(20.0).highlight("#FFFF00").build(),
            TextRunBuilder::new(" · ").font_size(20.0).build(),
            TextRunBuilder::new("Cyan highlight").font_size(20.0).highlight("#00FFFF").build(),
        ];
        s.add_text_runs(runs, TextOptionsBuilder::new().x(0.5).y(1.0).w(9.0).h(0.8).build());

        // Underline variants
        let underlines = [
            ("Single (sng)", "sng"),
            ("Double (dbl)", "dbl"),
            ("Dash (dash)", "dash"),
            ("Dotted", "dotted"),
            ("Heavy (heavy)", "heavy"),
            ("Wavy (wavy)", "wavy"),
            ("Wavy Double (wavyDbl)", "wavyDbl"),
        ];
        for (i, (label, style)) in underlines.iter().enumerate() {
            let col = i % 2;
            let row = i / 2;
            let x = 0.5 + col as f64 * 4.8;
            let y = 1.9 + row as f64 * 0.55;
            let runs = vec![TextRunBuilder::new(*label).font_size(18.0).underline(*style).build()];
            s.add_text_runs(runs, TextOptionsBuilder::new().x(x).y(y).w(4.5).h(0.5).build());
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Text Glow + Outline
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#1A1A2E");
        heading(s, "Text: Glow & Outline");

        // Glow — various radii and colors
        let glows = [
            ("Blue glow", "#4472C4", "#4472C4", 0.7_f64, 8.0_f64),
            ("Red glow", "#FF0000", "#FF6B6B", 0.8, 12.0),
            ("Green glow", "#FFFFFF", "#00FF88", 0.6, 6.0),
        ];
        for (i, (label, text_color, glow_color, opacity, size)) in glows.iter().enumerate() {
            let runs = vec![
                TextRunBuilder::new(*label)
                    .font_size(36.0).bold().color(*text_color)
                    .glow(GlowProps { size: *size, color: glow_color.to_string(), opacity: *opacity })
                    .build(),
            ];
            s.add_text_runs(
                runs,
                TextOptionsBuilder::new()
                    .x(0.5).y(1.0 + i as f64 * 1.1).w(9.0).h(0.9)
                    .align(AlignH::Center)
                    .build(),
            );
        }

        // Outline
        let runs = vec![
            TextRunBuilder::new("Outlined text (white fill, black stroke)")
                .font_size(28.0).bold().color("#FFFFFF")
                .outline(TextOutlineProps { color: "#000000".to_string(), size: 1.5 })
                .build(),
        ];
        s.add_text_runs(
            runs,
            TextOptionsBuilder::new()
                .x(0.5).y(4.4).w(9.0).h(0.9)
                .align(AlignH::Center)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Soft Break + Text Direction + Tab Stops
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Text: Soft Break · Direction · Tab Stops");

        // Soft breaks — all runs in one paragraph, separated by <a:br>
        let runs = vec![
            TextRunBuilder::new("Line 1 — first line of paragraph").font_size(16.0).build(),
            TextRunBuilder::new("Line 2 — soft break before this (same paragraph)")
                .font_size(16.0).color("#4472C4").soft_break_before().build(),
            TextRunBuilder::new("Line 3 — another soft break")
                .font_size(16.0).italic().soft_break_before().build(),
        ];
        s.add_text_runs(
            runs,
            TextOptionsBuilder::new().x(0.5).y(0.9).w(6.5).h(1.8).build(),
        );

        // Vertical text direction
        s.add_text(
            "VERTICAL",
            TextOptionsBuilder::new()
                .x(7.5).y(0.9).w(0.8).h(3.5)
                .font_size(16.0).bold().color("#C00000")
                .text_direction("vert")
                .build(),
        );
        s.add_text(
            "↑ vert270",
            TextOptionsBuilder::new()
                .x(8.5).y(0.9).w(0.9).h(3.5)
                .font_size(14.0).color("#4472C4")
                .text_direction("vert270")
                .build(),
        );

        // Tab stops
        s.add_text(
            "Name\tRole\tScore",
            TextOptionsBuilder::new()
                .x(0.5).y(2.9).w(9.0).h(0.5)
                .font_size(16.0).bold()
                .tab_stops(vec![
                    TabStop { pos_inches: 3.0, align: "l".to_string() },
                    TabStop { pos_inches: 6.5, align: "r".to_string() },
                ])
                .build(),
        );
        for (name, role, score) in &[("Alice", "Engineer", "98"), ("Bob", "Designer", "87"), ("Carol", "PM", "92")] {
            let row_y = 3.4 + match *name { "Alice" => 0.0, "Bob" => 0.45, _ => 0.9 };
            s.add_text(
                &format!("{name}\t{role}\t{score}"),
                TextOptionsBuilder::new()
                    .x(0.5).y(row_y).w(9.0).h(0.4)
                    .font_size(15.0)
                    .tab_stops(vec![
                        TabStop { pos_inches: 3.0, align: "l".to_string() },
                        TabStop { pos_inches: 6.5, align: "r".to_string() },
                    ])
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Shape Hyperlink + Shadow rotateWithShape
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Shapes: Hyperlink · Shadow rotateWithShape");

        // Shape with external URL hyperlink
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .x(0.5).y(1.0).w(4.0).h(1.0)
                .fill_color("#4472C4")
                .hyperlink(HyperlinkProps {
                    r_id: 0,
                    slide: None,
                    url: Some("https://github.com/gitbrent/PptxGenJS".to_string()),
                    tooltip: Some("Open PptxGenJS on GitHub".to_string()),
                    action: None,
                })
                .build(),
        );
        s.add_text(
            "▶  Click: GitHub link (shape hyperlink)",
            TextOptionsBuilder::new()
                .x(0.5).y(1.0).w(4.0).h(1.0)
                .font_size(14.0).bold().color("#FFFFFF")
                .valign(AlignV::Middle)
                .align(AlignH::Center)
                .build(),
        );

        // Shape with slide-jump hyperlink
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .x(5.0).y(1.0).w(4.0).h(1.0)
                .fill_color("#ED7D31")
                .hyperlink(HyperlinkProps {
                    r_id: 0,
                    slide: Some(1),
                    url: None,
                    tooltip: Some("Jump to slide 1".to_string()),
                    action: None,
                })
                .build(),
        );
        s.add_text(
            "▶  Click: Jump to Slide 1 (slide hyperlink)",
            TextOptionsBuilder::new()
                .x(5.0).y(1.0).w(4.0).h(1.0)
                .font_size(14.0).bold().color("#FFFFFF")
                .valign(AlignV::Middle)
                .align(AlignH::Center)
                .build(),
        );

        // Shadow: rotateWithShape = true vs false
        use deckmint::{ShadowProps, ShadowType};
        let shadow_on = ShadowProps {
            shadow_type: ShadowType::Outer,
            blur: Some(6.0),
            offset: Some(4.0),
            angle: Some(45.0),
            color: Some("#000000".to_string()),
            opacity: Some(0.5),
            rotate_with_shape: true,
        };
        let shadow_off = ShadowProps {
            rotate_with_shape: false,
            ..shadow_on.clone()
        };

        s.add_shape(
            ShapeType::Pentagon,
            ShapeOptionsBuilder::new()
                .x(0.5).y(2.5).w(3.0).h(2.0)
                .fill_color("#70AD47")
                .rotate(30.0)
                .shadow(shadow_on)
                .build(),
        );
        s.add_text(
            "rotateWithShape=true\n(shadow rotates with shape)",
            TextOptionsBuilder::new()
                .x(0.5).y(4.6).w(3.5).h(0.8)
                .font_size(12.0).align(AlignH::Center)
                .build(),
        );

        s.add_shape(
            ShapeType::Pentagon,
            ShapeOptionsBuilder::new()
                .x(5.5).y(2.5).w(3.0).h(2.0)
                .fill_color("#ED7D31")
                .rotate(30.0)
                .shadow(shadow_off)
                .build(),
        );
        s.add_text(
            "rotateWithShape=false\n(shadow stays fixed)",
            TextOptionsBuilder::new()
                .x(5.5).y(4.6).w(3.5).h(0.8)
                .font_size(12.0).align(AlignH::Center)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 5: Table — per-row heights + cell hyperlinks
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Tables: Per-Row Heights · Cell Hyperlinks");

        let mut header = TableCell::new("Feature");
        header.options.bold = Some(true);
        header.options.fill = Some("#2E4057".to_string());
        header.options.color = Some("#FFFFFF".to_string());
        let mut h2 = TableCell::new("Value");
        h2.options.bold = Some(true);
        h2.options.fill = Some("#2E4057".to_string());
        h2.options.color = Some("#FFFFFF".to_string());
        let mut h3 = TableCell::new("Link");
        h3.options.bold = Some(true);
        h3.options.fill = Some("#2E4057".to_string());
        h3.options.color = Some("#FFFFFF".to_string());

        // Row with cell hyperlink
        let mut link_cell = TableCell::new("Click me");
        link_cell.options.color = Some("#4472C4".to_string());
        link_cell.options.underline = Some(true);
        link_cell.options.hyperlink = Some(HyperlinkProps {
            r_id: 0,
            slide: None,
            url: Some("https://github.com/gitbrent/PptxGenJS".to_string()),
            tooltip: Some("Open PptxGenJS".to_string()),
            action: None,
        });

        let rows = vec![
            vec![header, h2, h3],
            vec![TableCell::new("Row height 0.6\""), TableCell::new("Tall row"), TableCell::new("—")],
            vec![TableCell::new("Row height 0.3\""), TableCell::new("Normal row"), TableCell::new("—")],
            vec![TableCell::new("Cell hyperlink"), TableCell::new("Underlined blue text"), link_cell],
            vec![TableCell::new("Row height 0.8\""), TableCell::new("Extra tall"), TableCell::new("—")],
        ];

        s.add_table(
            rows,
            TableOptionsBuilder::new()
                .x(0.5).y(0.9).w(9.0).h(5.5)
                .col_w(vec![3.0, 3.5, 2.5])
                .row_h(vec![0.5, 0.6, 0.3, 0.4, 0.8])
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 6: Background image
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        // The 1×1 red PNG tiles/stretches as a background fill
        s.set_background_image(RED_1PX_PNG.to_vec(), "png");
        s.add_text(
            "Background Image",
            TextOptionsBuilder::new()
                .x(0.5).y(2.5).w(9.0).h(1.2)
                .font_size(40.0).bold()
                .align(AlignH::Center).color("#FFFFFF")
                .build(),
        );
        s.add_text(
            "Slide background set from raw image bytes via set_background_image()",
            TextOptionsBuilder::new()
                .x(0.5).y(3.8).w(9.0).h(0.8)
                .font_size(16.0)
                .align(AlignH::Center).color("#FFFFFF")
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 7: Slide master objects visible on this slide
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Slide Master: Bottom Stripe from Master");
        s.add_text(
            "The dark stripe at the bottom of every slide in this deck\n\
is defined on the slide master via define_master(). It appears on\n\
all slides without being added individually.",
            TextOptionsBuilder::new()
                .x(0.5).y(1.0).w(9.0).h(2.5)
                .font_size(18.0)
                .build(),
        );
        s.add_text(
            "Custom slide layouts can also be added with define_layout(name, w, h)\n\
which supports custom slide dimensions.",
            TextOptionsBuilder::new()
                .x(0.5).y(3.8).w(9.0).h(1.5)
                .font_size(16.0).color("#555555")
                .build(),
        );
    }

    // ── Slide: Animations — entrance / exit ──────────────
    {
        let s = pres.add_slide();
        s.add_text("Animations: Entrance & Exit", TextOptionsBuilder::new().x(0.5).y(0.2).w(9.0).h(0.6).font_size(26.0).bold().build());

        let items: &[(&str, fn() -> AnimationEffect)] = &[
            ("Appear (entrance)",       || AnimationEffect::appear()),
            ("Disappear (exit)",        || AnimationEffect::disappear()),
            ("Fade In (entrance)",      || AnimationEffect::fade_in()),
            ("Fade Out (exit)",         || AnimationEffect::fade_out()),
            ("Fly In from Left",        || AnimationEffect::fly_in(Direction::Left)),
            ("Fly Out to Right",        || AnimationEffect::fly_out(Direction::Right)),
            ("Fly In from Bottom",      || AnimationEffect::fly_in(Direction::Down)),
            ("Wipe In from Left",       || AnimationEffect::wipe_in(Direction::Left)),
            ("Zoom In (entrance)",      || AnimationEffect::zoom_in()),
            ("Zoom Out (exit)",         || AnimationEffect::zoom_out()),
        ];
        for (idx, (label, make)) in items.iter().enumerate() {
            let col = (idx % 2) as f64;
            let row = (idx / 2) as f64;
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .x(0.4 + col * 4.8).y(1.0 + row * 0.7).w(4.5).h(0.6)
                    .font_size(16.0)
                    .animation(make())
                    .build(),
            );
        }
    }

    // ── Slide: Animations — emphasis & split ─────────────
    {
        let s = pres.add_slide();
        s.add_text("Animations: Emphasis & Split", TextOptionsBuilder::new().x(0.5).y(0.2).w(9.0).h(0.6).font_size(26.0).bold().build());

        s.add_text("Spin 360°", TextOptionsBuilder::new()
            .x(0.5).y(1.1).w(4.0).h(0.6).font_size(18.0)
            .animation(AnimationEffect::spin(360.0)).build());
        s.add_text("Spin 180° (half)", TextOptionsBuilder::new()
            .x(5.0).y(1.1).w(4.0).h(0.6).font_size(18.0)
            .animation(AnimationEffect::spin(180.0)).build());
        s.add_text("Pulse", TextOptionsBuilder::new()
            .x(0.5).y(1.9).w(4.0).h(0.6).font_size(18.0)
            .animation(AnimationEffect::pulse()).build());
        s.add_text("Grow/Shrink ×1.5", TextOptionsBuilder::new()
            .x(5.0).y(1.9).w(4.0).h(0.6).font_size(18.0)
            .animation(AnimationEffect::grow_shrink(1.5)).build());
        s.add_text("Split In (Horizontal)", TextOptionsBuilder::new()
            .x(0.5).y(2.7).w(4.0).h(0.6).font_size(18.0)
            .animation(AnimationEffect::split_in(SplitOrientation::Horizontal)).build());
        s.add_text("Split Out (Vertical)", TextOptionsBuilder::new()
            .x(5.0).y(2.7).w(4.0).h(0.6).font_size(18.0)
            .animation(AnimationEffect::split_out(SplitOrientation::Vertical)).build());
        s.add_text("Wipe In from Top", TextOptionsBuilder::new()
            .x(0.5).y(3.5).w(4.0).h(0.6).font_size(18.0)
            .animation(AnimationEffect::wipe_in(Direction::Up)).build());
        s.add_text("Fly In from Top", TextOptionsBuilder::new()
            .x(5.0).y(3.5).w(4.0).h(0.6).font_size(18.0)
            .animation(AnimationEffect::fly_in(Direction::Up)).build());
    }

    // ══════════════════════════════════════════════════════════
    // Slide: GridLayout — 3 equal columns
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Layout: GridLayout — 3 Equal Columns");

        let grid = GridLayoutBuilder::cols_n(3, 0.2)
            .origin(0.3, 0.9)
            .container(9.4, 4.6)
            .build();

        let colors = ["#4472C4", "#ED7D31", "#70AD47"];
        let labels = ["Column 1\n(Fr 1)", "Column 2\n(Fr 1)", "Column 3\n(Fr 1)"];
        for col in 0..3 {
            let r = grid.cell(col, 0);
            s.add_shape(ShapeType::Rect,
                ShapeOptionsBuilder::new()
                    .x(r.x).y(r.y).w(r.w).h(r.h)
                    .fill_color(colors[col]).build());
            s.add_text(labels[col],
                TextOptionsBuilder::new()
                    .x(r.x).y(r.y).w(r.w).h(r.h)
                    .font_size(20.0).bold().color("#FFFFFF")
                    .align(AlignH::Center).valign(AlignV::Middle).build());
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide: GridLayout — Mixed tracks (sidebar + content grid)
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Layout: GridLayout — Mixed Tracks & Spans");

        // Outer grid: fixed header row + flexible content + fixed footer
        let outer = GridLayoutBuilder::new()
            .cols(vec![GridTrack::Fr(1.0)])
            .rows(vec![
                GridTrack::Inches(0.55),  // header (occupied by heading())
                GridTrack::Fr(1.0),        // content
                GridTrack::Inches(0.25),   // footer bar
            ])
            .row_gap(0.05)
            .origin(0.3, 0.85)
            .container(9.4, 4.65)
            .build();

        // Inner content grid: [fixed sidebar | 2-col main area] × 2 rows
        let content_cell = outer.cell(0, 1);
        let inner = GridLayoutBuilder::new()
            .cols(vec![GridTrack::Inches(2.2), GridTrack::Fr(1.0), GridTrack::Fr(1.0)])
            .rows(vec![GridTrack::Fr(1.0), GridTrack::Fr(1.0)])
            .col_gap(0.15)
            .row_gap(0.12)
            .origin(content_cell.x, content_cell.y)
            .container(content_cell.w, content_cell.h)
            .build();

        // Sidebar spans both rows
        let sidebar = inner.span(0, 0, 1, 2);
        s.add_shape(ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .x(sidebar.x).y(sidebar.y).w(sidebar.w).h(sidebar.h)
                .fill_color("#2E4057").build());
        s.add_text("Sidebar\n(fixed 2.2\")",
            TextOptionsBuilder::new()
                .x(sidebar.x).y(sidebar.y).w(sidebar.w).h(sidebar.h)
                .font_size(14.0).bold().color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle).build());

        // 4 content cells
        let cell_colors = [("#4472C4","Top Left"), ("#ED7D31","Top Right"),
                           ("#70AD47","Bottom Left"), ("#9E480E","Bottom Right")];
        for (idx, (color, label)) in cell_colors.iter().enumerate() {
            let col = 1 + (idx % 2);
            let row = idx / 2;
            let r = inner.cell(col, row);
            s.add_shape(ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .x(r.x).y(r.y).w(r.w).h(r.h)
                    .fill_color(*color).rect_radius(0.05).build());
            s.add_text(*label,
                TextOptionsBuilder::new()
                    .x(r.x).y(r.y).w(r.w).h(r.h)
                    .font_size(13.0).bold().color("#FFFFFF")
                    .align(AlignH::Center).valign(AlignV::Middle).build());
        }

        // Footer: centered label using center_in helper
        let footer = outer.cell(0, 2);
        let centered_label = center_in(&footer, 4.0, footer.h);
        s.add_text("Footer — auto-positioned with center_in()",
            TextOptionsBuilder::new()
                .x(centered_label.x).y(centered_label.y)
                .w(centered_label.w).h(centered_label.h)
                .font_size(10.0).italic().color("#888888")
                .align(AlignH::Center).build());
    }

    // ══════════════════════════════════════════════════════════
    // Slide: Gradient Fills
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Gradient Fills");

        let grid = GridLayoutBuilder::grid_n_m(3, 2, 0.18)
            .origin(0.3, 0.95)
            .container(9.4, 4.55)
            .build();

        // Row 0 — linear gradients at different angles
        let r = grid.cell(0, 0);
        s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h)
            .gradient_fill(GradientFill::two_color(0.0, "#4472C4", "#ED7D31"))
            .build());
        s.add_text("Linear 0°\n(left → right)", TextOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h).font_size(13.0).bold().color("#FFFFFF")
            .align(AlignH::Center).valign(AlignV::Middle).build());

        let r = grid.cell(1, 0);
        s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h)
            .gradient_fill(GradientFill::two_color(90.0, "#4472C4", "#ED7D31"))
            .build());
        s.add_text("Linear 90°\n(top → bottom)", TextOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h).font_size(13.0).bold().color("#FFFFFF")
            .align(AlignH::Center).valign(AlignV::Middle).build());

        let r = grid.cell(2, 0);
        s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h)
            .gradient_fill(GradientFill::two_color(45.0, "#4472C4", "#ED7D31"))
            .build());
        s.add_text("Linear 45°\n(diagonal)", TextOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h).font_size(13.0).bold().color("#FFFFFF")
            .align(AlignH::Center).valign(AlignV::Middle).build());

        // Row 1 — multi-stop, radial, text-box gradient
        let r = grid.cell(0, 1);
        s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h)
            .gradient_fill(GradientFill::linear(90.0, vec![
                GradientStop::new("#FF0000", 0.0),
                GradientStop::new("#FFFF00", 50.0),
                GradientStop::new("#00B050", 100.0),
            ]))
            .build());
        s.add_text("Multi-stop\n(3 colours)", TextOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h).font_size(13.0).bold().color("#FFFFFF")
            .align(AlignH::Center).valign(AlignV::Middle).build());

        let r = grid.cell(1, 1);
        s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h)
            .gradient_fill(GradientFill::radial_two_color("#FFFFFF", "#4472C4"))
            .build());
        s.add_text("Radial\n(white → blue)", TextOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h).font_size(13.0).bold().color("#1F3864")
            .align(AlignH::Center).valign(AlignV::Middle).build());

        let r = grid.cell(2, 1);
        s.add_text("Gradient\ntext box\nbackground", TextOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h).font_size(13.0).bold().color("#FFFFFF")
            .align(AlignH::Center).valign(AlignV::Middle)
            .gradient_fill(GradientFill::two_color(135.0, "#7030A0", "#00B0F0"))
            .build());
    }

    // ── Slide: New Entrance Animations (Basic filter-based) ─────
    {
        let s = pres.add_slide();
        heading(s, "New Entrance Animations — Basic (Filter-Based)");

        let grid = GridLayoutBuilder::new()
            .cols(vec![GridTrack::fr(1.0); 4])
            .rows(vec![GridTrack::inches(0.6), GridTrack::fr(1.0), GridTrack::fr(1.0), GridTrack::fr(1.0)])
            .gap(0.15)
            .origin(0.3, 0.85)
            .container(9.4, 4.6)
            .build();

        let anims: &[(&str, AnimationEffect, &str)] = &[
            ("Blinds H",    AnimationEffect::blinds_in(SplitOrientation::Horizontal), "#4472C4"),
            ("Blinds V",    AnimationEffect::blinds_in(SplitOrientation::Vertical),   "#ED7D31"),
            ("Checker\n(Across)", AnimationEffect::checkerboard_in(CheckerboardDir::Across), "#70AD47"),
            ("Checker\n(Down)",   AnimationEffect::checkerboard_in(CheckerboardDir::Down),   "#FFC000"),
            ("Dissolve In", AnimationEffect::dissolve_in(),                           "#5B9BD5"),
            ("Peek In",     AnimationEffect::peek_in(Direction::Down),                "#C55A11"),
            ("Rand Bars H", AnimationEffect::random_bars_in(SplitOrientation::Horizontal), "#7030A0"),
            ("Rand Bars V", AnimationEffect::random_bars_in(SplitOrientation::Vertical),   "#00B0F0"),
            ("Shape Box",   AnimationEffect::shape_in(ShapeVariant::Box),             "#4472C4"),
            ("Shape Circle",AnimationEffect::shape_in(ShapeVariant::Circle),          "#ED7D31"),
            ("Strips LD",   AnimationEffect::strips_in(StripDir::LeftDown),           "#70AD47"),
            ("Wedge",       AnimationEffect::wedge_in(),                              "#FFC000"),
        ];

        let cols = 4;
        for (i, (label, anim, color)) in anims.iter().enumerate() {
            let col = i % cols;
            let row = i / cols + 1;
            let r = grid.cell(col, row);
            s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h)
                .fill_color(*color)
                .animation(anim.clone())
                .build());
            s.add_text(*label, TextOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h).font_size(11.0).bold().color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle).build());
        }

        // Wheel with 3 spokes
        let r = grid.cell(3, 3);
        s.add_shape(ShapeType::Ellipse, ShapeOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h)
            .fill_color("#FF0000")
            .animation(AnimationEffect::wheel_in(3))
            .build());
        s.add_text("Wheel\n(3 spokes)", TextOptionsBuilder::new()
            .x(r.x).y(r.y).w(r.w).h(r.h).font_size(11.0).bold().color("#FFFFFF")
            .align(AlignH::Center).valign(AlignV::Middle).build());
    }

    // ── Slide: New Entrance Animations (Subtle/Moderate) ────────
    {
        let s = pres.add_slide();
        heading(s, "New Entrance Animations — Subtle & Moderate");

        let grid = GridLayoutBuilder::new()
            .cols(vec![GridTrack::fr(1.0); 4])
            .rows(vec![GridTrack::inches(0.6), GridTrack::fr(1.0), GridTrack::fr(1.0)])
            .gap(0.15)
            .origin(0.3, 0.85)
            .container(9.4, 4.6)
            .build();

        let anims: &[(&str, AnimationEffect, &str)] = &[
            ("Expand",         AnimationEffect::expand_in(),                         "#4472C4"),
            ("Swivel",         AnimationEffect::swivel_in(),                         "#ED7D31"),
            ("Basic Zoom",     AnimationEffect::basic_zoom_in(),                     "#70AD47"),
            ("Stretch H",      AnimationEffect::stretch_in(Direction::Left),         "#FFC000"),
            ("Centre Revolve", AnimationEffect::centre_revolve_in(),                 "#7030A0"),
            ("Float In ↑",     AnimationEffect::float_in(Direction::Up),             "#00B0F0"),
            ("Grow Turn",      AnimationEffect::grow_turn_in(),                      "#C55A11"),
            ("Rise Up",        AnimationEffect::rise_up_in(),                        "#4472C4"),
        ];

        let cols = 4;
        for (i, (label, anim, color)) in anims.iter().enumerate() {
            let col = i % cols;
            let row = i / cols + 1;
            let r = grid.cell(col, row);
            s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h)
                .fill_color(*color)
                .animation(anim.clone())
                .build());
            s.add_text(*label, TextOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h).font_size(11.0).bold().color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle).build());
        }
    }

    // ── Slide: New Entrance Animations (Exciting) ───────────────
    {
        let s = pres.add_slide();
        heading(s, "New Entrance Animations — Exciting");

        let grid = GridLayoutBuilder::new()
            .cols(vec![GridTrack::fr(1.0); 4])
            .rows(vec![GridTrack::inches(0.6), GridTrack::fr(1.0), GridTrack::fr(1.0), GridTrack::fr(1.0)])
            .gap(0.15)
            .origin(0.3, 0.85)
            .container(9.4, 4.6)
            .build();

        let anims: &[(&str, AnimationEffect, &str)] = &[
            ("Boomerang",     AnimationEffect::boomerang_in(),    "#4472C4"),
            ("Bounce",        AnimationEffect::bounce_in(),       "#ED7D31"),
            ("Credits",       AnimationEffect::credits_in(),      "#70AD47"),
            ("Curve Up",      AnimationEffect::curve_up_in(),     "#FFC000"),
            ("Drop",          AnimationEffect::drop_in(),         "#7030A0"),
            ("Flip",          AnimationEffect::flip_in(),         "#00B0F0"),
            ("Pinwheel",      AnimationEffect::pinwheel_in(),     "#C55A11"),
            ("Spiral In",     AnimationEffect::spiral_in(),       "#4472C4"),
            ("Basic Swivel",  AnimationEffect::basic_swivel_in(), "#ED7D31"),
            ("Whip",          AnimationEffect::whip_in(),         "#70AD47"),
            ("Spinner",       AnimationEffect::spinner_in(),      "#FFC000"),
        ];

        let cols = 4;
        for (i, (label, anim, color)) in anims.iter().enumerate() {
            let col = i % cols;
            let row = i / cols + 1;
            let r = grid.cell(col, row);
            s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h)
                .fill_color(*color)
                .animation(anim.clone())
                .build());
            s.add_text(*label, TextOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h).font_size(11.0).bold().color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle).build());
        }
    }

    // ── Exit animations — Basic / Subtle ────────────────────────────────────
    {
        let s = pres.add_slide();
        heading(s, "Exit Animations — Basic & Subtle");
        let grid = GridLayoutBuilder::grid_n_m(4, 4, 0.08)
            .origin(0.3, 0.85).container(9.4, 4.7).build();

        let anims: &[(&str, AnimationEffect, &str)] = &[
            ("Blinds H",      AnimationEffect::blinds_out(SplitOrientation::Horizontal), "#4472C4"),
            ("Blinds V",      AnimationEffect::blinds_out(SplitOrientation::Vertical),   "#4472C4"),
            ("Checkerboard",  AnimationEffect::checkerboard_out(CheckerboardDir::Across), "#ED7D31"),
            ("Dissolve Out",  AnimationEffect::dissolve_out(),                            "#A9D18E"),
            ("Peek Out",      AnimationEffect::peek_out(Direction::Left),                 "#FF0000"),
            ("Random Bars H", AnimationEffect::random_bars_out(SplitOrientation::Horizontal), "#70AD47"),
            ("Random Bars V", AnimationEffect::random_bars_out(SplitOrientation::Vertical),   "#70AD47"),
            ("Shape Box",     AnimationEffect::shape_out(ShapeVariant::Box),              "#FFC000"),
            ("Strips",        AnimationEffect::strips_out(StripDir::LeftDown),            "#9DC3E6"),
            ("Wedge",         AnimationEffect::wedge_out(),                               "#FF7F7F"),
            ("Wheel 2",       AnimationEffect::wheel_out(2),                             "#C5A5C5"),
            ("Wipe Left",     AnimationEffect::wipe_out(Direction::Left),                 "#4472C4"),
            ("Contract",      AnimationEffect::contract_out(),                            "#ED7D31"),
            ("Swivel",        AnimationEffect::swivel_out(),                              "#70AD47"),
            ("Split H",       AnimationEffect::split_out(SplitOrientation::Horizontal),  "#FF0000"),
            ("Zoom Out",      AnimationEffect::zoom_out(),                               "#FFC000"),
        ];

        let cols = 4;
        for (i, (label, anim, color)) in anims.iter().enumerate() {
            let col = i % cols;
            let row = i / cols;
            let r = grid.cell(col, row);
            s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h)
                .fill_color(*color)
                .animation(anim.clone())
                .build());
            s.add_text(*label, TextOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h).font_size(11.0).bold().color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle).build());
        }
    }

    // ── Exit animations — Moderate & Exciting ───────────────────────────────
    {
        let s = pres.add_slide();
        heading(s, "Exit Animations — Moderate & Exciting");
        let grid = GridLayoutBuilder::grid_n_m(4, 4, 0.08)
            .origin(0.3, 0.85).container(9.4, 4.7).build();

        let anims: &[(&str, AnimationEffect, &str)] = &[
            ("Centre Revolve", AnimationEffect::centre_revolve_out(),              "#4472C4"),
            ("Collapse",       AnimationEffect::collapse_out(),                    "#ED7D31"),
            ("Float Out",      AnimationEffect::float_out(Direction::Up),          "#A9D18E"),
            ("Shrink Turn",    AnimationEffect::shrink_turn_out(),                 "#FF0000"),
            ("Sink Down",      AnimationEffect::sink_down_out(),                   "#70AD47"),
            ("Spinner",        AnimationEffect::spinner_out(),                     "#FFC000"),
            ("Basic Zoom",     AnimationEffect::basic_zoom_out(),                  "#9DC3E6"),
            ("Stretchy H",     AnimationEffect::stretchy_out(Direction::Left),     "#FF7F7F"),
            ("Boomerang",      AnimationEffect::boomerang_out(),                   "#C5A5C5"),
            ("Bounce Out",     AnimationEffect::bounce_out(),                      "#4472C4"),
            ("Credits Out",    AnimationEffect::credits_out(),                     "#ED7D31"),
            ("Curve Down",     AnimationEffect::curve_down_out(),                  "#70AD47"),
            ("Drop Out",       AnimationEffect::drop_out(),                        "#FF0000"),
            ("Flip Out",       AnimationEffect::flip_out(),                        "#FFC000"),
            ("Pinwheel Out",   AnimationEffect::pinwheel_out(),                    "#9DC3E6"),
            ("Whip Out",       AnimationEffect::whip_out(),                        "#A9D18E"),
        ];

        let cols = 4;
        for (i, (label, anim, color)) in anims.iter().enumerate() {
            let col = i % cols;
            let row = i / cols;
            let r = grid.cell(col, row);
            s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h)
                .fill_color(*color)
                .animation(anim.clone())
                .build());
            s.add_text(*label, TextOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h).font_size(11.0).bold().color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle).build());
        }
    }

    // ── Emphasis animations — Basic / Subtle ────────────────────────────────
    {
        let s = pres.add_slide();
        heading(s, "Emphasis — Basic & Subtle");
        let grid = GridLayoutBuilder::grid_n_m(4, 4, 0.08)
            .origin(0.3, 0.85).container(9.4, 4.7).build();

        // on_text=true → animation targets the text label (needed for font/text-property effects)
        // on_text=false → animation targets the shape fill/geometry
        let anims: &[(&str, AnimationEffect, &str, bool)] = &[
            ("Fill Colour",   AnimationEffect::fill_color("#FF0000"),          "#4472C4", false),
            ("Font Colour",   AnimationEffect::font_color("#FFFF00"),          "#ED7D31", true),
            ("Line Colour",   AnimationEffect::line_color("#00FF00"),          "#70AD47", false),
            ("Transparency",  AnimationEffect::transparency(0.2),             "#FF0000", false),
            ("Bold Flash",    AnimationEffect::bold_flash(),                  "#4472C4", true),
            ("Brush Colour",  AnimationEffect::brush_color("#FF8000"),         "#ED7D31", false),
            ("Comp. Colour",  AnimationEffect::complementary_color(),         "#70AD47", false),
            ("Comp. Colour2", AnimationEffect::complementary_color2(),        "#FFC000", false),
            ("Contrast Clr",  AnimationEffect::contrasting_color(),           "#9DC3E6", false),
            ("Darken",        AnimationEffect::darken(),                      "#FF7F7F", false),
            ("Desaturate",    AnimationEffect::desaturate(),                  "#C5A5C5", false),
            ("Lighten",       AnimationEffect::lighten(),                     "#4472C4", false),
            ("Object Colour", AnimationEffect::object_color("#70AD47"),        "#ED7D31", false),
            ("Underline",     AnimationEffect::underline(),                   "#FF0000", true),
            ("Spin 360°",     AnimationEffect::spin(360.0),                   "#70AD47", false),
            ("Pulse",         AnimationEffect::pulse(),                       "#FFC000", false),
        ];

        let cols = 4;
        for (i, (label, anim, color, on_text)) in anims.iter().enumerate() {
            let col = i % cols;
            let row = i / cols;
            let r = grid.cell(col, row);
            if *on_text {
                // Shape provides the background; text element carries the animation
                s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
                    .x(r.x).y(r.y).w(r.w).h(r.h).fill_color(*color).build());
                s.add_text(*label, TextOptionsBuilder::new()
                    .x(r.x).y(r.y).w(r.w).h(r.h).font_size(10.5).bold().color("#FFFFFF")
                    .align(AlignH::Center).valign(AlignV::Middle)
                    .animation(anim.clone()).build());
            } else {
                s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
                    .x(r.x).y(r.y).w(r.w).h(r.h).fill_color(*color).animation(anim.clone()).build());
                s.add_text(*label, TextOptionsBuilder::new()
                    .x(r.x).y(r.y).w(r.w).h(r.h).font_size(10.5).bold().color("#FFFFFF")
                    .align(AlignH::Center).valign(AlignV::Middle).build());
            }
        }
    }

    // ── Emphasis animations — Moderate / Exciting ────────────────────────────
    {
        let s = pres.add_slide();
        heading(s, "Emphasis — Moderate & Exciting");
        let grid = GridLayoutBuilder::grid_n_m(4, 2, 0.08)
            .origin(0.3, 0.85).container(9.4, 4.7).build();

        let anims: &[(&str, AnimationEffect, &str)] = &[
            ("Colour Pulse",    AnimationEffect::color_pulse("#ED7D31"),        "#4472C4"),
            ("Grow w/ Colour",  AnimationEffect::grow_with_color("#FF0000"),    "#70AD47"),
            ("Shimmer",         AnimationEffect::shimmer(),                    "#FFC000"),
            ("Teeter",          AnimationEffect::teeter(),                     "#9DC3E6"),
            ("Blink",           AnimationEffect::blink(),                      "#FF7F7F"),
            ("Bold Reveal",     AnimationEffect::bold_reveal(),                "#C5A5C5"),
            ("Wave",            AnimationEffect::wave(),                       "#4472C4"),
            ("Grow/Shrink",     AnimationEffect::grow_shrink(1.5),             "#ED7D31"),
        ];

        let cols = 4;
        for (i, (label, anim, color)) in anims.iter().enumerate() {
            let col = i % cols;
            let row = i / cols;
            let r = grid.cell(col, row);
            s.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h)
                .fill_color(*color)
                .animation(anim.clone())
                .build());
            s.add_text(*label, TextOptionsBuilder::new()
                .x(r.x).y(r.y).w(r.w).h(r.h).font_size(11.0).bold().color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle).build());
        }
    }

    // ── Sub-paragraph text targeting demo ────────────────────────────────────
    {
        let s = pres.add_slide();
        heading(s, "Font Colour — Sub-paragraph targeting");

        // "The quick brown fox" — FontColor on "quick" (chars 4..8) flashes yellow
        // "jumps over the lazy dog" — FontColor on "over" (chars 6..9) flashes red
        // Each is a separate text box so two independent animations can fire.
        let demos: &[(&str, u32, u32, &str, &str)] = &[
            // (text, st_idx, end_idx, target_word, color)
            ("The quick brown fox",    4, 8, "\"quick\"",  "#FFFF00"),
            ("jumps over the lazy dog", 6, 9, "\"over\"",  "#FF4444"),
        ];
        for (row_idx, (text, st, end, word, clr)) in demos.iter().enumerate() {
            let y = 1.5 + row_idx as f64 * 2.0;
            // Instruction label
            s.add_text(
                &format!("Click to colour {} → #{} (chars {}–{})", word, clr, st, end),
                TextOptionsBuilder::new()
                    .x(0.5).y(y - 0.45).w(9.0).h(0.4).font_size(11.0).color("#666666").build()
            );
            // Animated text box — only the specified char range changes colour
            s.add_text(
                *text,
                TextOptionsBuilder::new()
                    .x(0.5).y(y).w(9.0).h(0.7)
                    .font_size(28.0).bold().color("#222222")
                    .animation(
                        AnimationEffect::font_color(clr)
                            .with_char_range(*st, *end)
                    )
                    .build()
            );
        }
        // Also show a whole-paragraph range example
        s.add_text(
            "Three paragraphs:\nFirst\nSecond\nThird",
            TextOptionsBuilder::new()
                .x(0.5).y(5.0).w(4.5).h(0.55).font_size(14.0).color("#444444").build()
        );
        s.add_text(
            "First\nSecond\nThird",
            TextOptionsBuilder::new()
                .x(5.2).y(5.0).w(4.3).h(0.55)
                .font_size(14.0).color("#222222")
                .animation(
                    AnimationEffect::font_color("#0070C0")
                        .with_para_range(1, 1)  // only "Second" (paragraph index 1)
                )
                .build()
        );
        s.add_text(
            "← click to colour only para 1 (\"Second\") blue",
            TextOptionsBuilder::new()
                .x(0.5).y(4.55).w(9.0).h(0.4).font_size(11.0).color("#666666").build()
        );
    }

    // ── Rich text — multiple colors within one paragraph ─────────────────────
    {
        let s = pres.add_slide();
        heading(s, "Rich Text — Per-run colors within a paragraph");

        // Row 1: simple sentence with two differently-colored words
        s.add_text_runs(
            vec![
                TextRunBuilder::new("The ").font_size(28.0).color("#222222").build(),
                TextRunBuilder::new("quick brown").font_size(28.0).color("#C45911").bold().build(),
                TextRunBuilder::new(" fox ").font_size(28.0).color("#222222").build(),
                TextRunBuilder::new("jumps").font_size(28.0).color("#2E75B6").bold().build(),
                TextRunBuilder::new(" over the lazy dog.").font_size(28.0).color("#222222").build(),
            ],
            TextOptionsBuilder::new().x(0.5).y(0.9).w(9.0).h(0.7).build(),
        );

        // Row 2: traffic-light colors for three clauses
        s.add_text_runs(
            vec![
                TextRunBuilder::new("Passed: ").font_size(22.0).color("#375623").bold().build(),
                TextRunBuilder::new("all tests green  ").font_size(22.0).color("#375623").build(),
                TextRunBuilder::new("  Warning: ").font_size(22.0).color("#7F6000").bold().build(),
                TextRunBuilder::new("deprecated API  ").font_size(22.0).color("#7F6000").build(),
                TextRunBuilder::new("  Error: ").font_size(22.0).color("#C00000").bold().build(),
                TextRunBuilder::new("build failed").font_size(22.0).color("#C00000").build(),
            ],
            TextOptionsBuilder::new().x(0.5).y(1.8).w(9.0).h(0.6).build(),
        );

        // Row 3: mixed sizes + italic monospace for an inline code snippet
        s.add_text_runs(
            vec![
                TextRunBuilder::new("Call ").font_size(20.0).color("#222222").build(),
                TextRunBuilder::new("slide.add_text_runs(runs, opts)")
                    .font_size(18.0).color("#7030A0").italic().font_face("Courier New").build(),
                TextRunBuilder::new(" for inline rich text.").font_size(20.0).color("#222222").build(),
            ],
            TextOptionsBuilder::new().x(0.5).y(2.6).w(9.0).h(0.5).build(),
        );

        // Row 4: break_line splits into separate paragraphs
        s.add_text_runs(
            vec![
                TextRunBuilder::new("First sentence stays on line one.")
                    .font_size(18.0).color("#1F3864").break_line().build(),
                TextRunBuilder::new("Second sentence is on its own line (")
                    .font_size(18.0).color("#C00000").build(),
                TextRunBuilder::new("new paragraph").font_size(18.0).color("#C00000").bold().build(),
                TextRunBuilder::new(").").font_size(18.0).color("#C00000").build(),
            ],
            TextOptionsBuilder::new().x(0.5).y(3.3).w(9.0).h(0.85).build(),
        );

        // Annotation box
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .x(0.3).y(4.35).w(9.4).h(1.5)
                .fill_color("#F2F2F2")
                .line_color("#BFBFBF").line_width(0.5)
                .build(),
        );
        s.add_text(
            "API: TextRunBuilder::new(text).color(hex).bold().italic().font_size(pt).break_line().build()\n\
             • Runs without .break_line() share a paragraph (same <a:p>)\n\
             • .break_line() on a run ends that paragraph and starts a new one",
            TextOptionsBuilder::new()
                .x(0.5).y(4.4).w(9.0).h(1.4)
                .font_size(11.0).color("#444444")
                .build(),
        );
    }

    // ── Per-sentence colour animation — separate text box per sentence ───────────
    // NOTE: <p:animClr style.color> with charRg applies to the WHOLE text body in
    // PowerPoint; charRg is ignored for the colour change itself.
    // Correct approach: one text box per animated sentence, each with its own
    // whole-box font_color animation.  Position them to look like one paragraph.
    {
        let s = pres.add_slide();
        heading(s, "Emphasis — Per-sentence colour animation (one box per sentence)");

        s.add_text(
            "Each sentence is a separate text box with its own font_color emphasis.",
            TextOptionsBuilder::new()
                .x(0.5).y(0.85).w(9.0).h(0.38).font_size(11.5).color("#666666").italic().build(),
        );

        // Five sentences — alternating orange/blue on successive clicks.
        // All at x=0.5, w=9.0; y spaced by 0.52" (one line at 22pt).
        let sentences: &[(&str, &str)] = &[
            ("The sun rose slowly over the hills.",          "#C45911"),
            ("Birds began to sing in the distance.",         "#2E75B6"),
            ("A gentle breeze stirred the leaves.",          "#C45911"),
            ("The day had finally begun.",                   "#2E75B6"),
            ("All was calm and still.",                      "#C45911"),
        ];
        let x = 0.5_f64;
        let y0 = 1.35_f64;
        let line_h = 0.52_f64;
        for (i, (text, color)) in sentences.iter().enumerate() {
            let y = y0 + i as f64 * line_h;
            s.add_text(
                *text,
                TextOptionsBuilder::new()
                    .x(x).y(y).w(9.0).h(line_h)
                    .font_size(22.0).color("#333333")
                    .animation(AnimationEffect::font_color(*color))
                    .build(),
            );
        }

        // Click guide
        s.add_text(
            "Clicks 1,3,5 → orange     Clicks 2,4 → blue     (each click = one sentence)",
            TextOptionsBuilder::new()
                .x(0.5).y(4.1).w(9.0).h(0.4).font_size(11.5).color("#666666").italic().build(),
        );

        // Annotation box
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .x(0.3).y(4.6).w(9.4).h(1.0)
                .fill_color("#F2F2F2").line_color("#BFBFBF").line_width(0.5).build(),
        );
        s.add_text(
            "for (text, color) in sentences {\n\
             \x20   slide.add_text(text, TextOptionsBuilder::new()\n\
             \x20       .x(0.5).y(y).w(9.0).h(0.52).font_size(22.0)\n\
             \x20       .animation(AnimationEffect::font_color(color)).build());\n\
             }",
            TextOptionsBuilder::new()
                .x(0.5).y(4.65).w(9.0).h(0.9)
                .font_size(10.0).color("#333333").font_face("Courier New").build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Milestone 1a/b — Image & Text rotate / flip
    // ══════════════════════════════════════════════════════════
    {
        use deckmint::objects::image::ImageOptionsBuilder;
        let s = pres.add_slide();
        heading(s, "New: Image & Text Rotate / Flip");

        // Image: normal (no transform)
        s.add_image(
            RED_1PX_PNG.to_vec(), "png",
            ImageOptionsBuilder::new().x(0.3).y(1.0).w(2.5).h(1.5).build().0,
        );
        s.add_text("image (normal)", TextOptionsBuilder::new().x(0.3).y(2.55).w(2.5).h(0.4)
            .font_size(11.0).align(AlignH::Center).build());

        // Image: rotate 30°
        s.add_image(
            RED_1PX_PNG.to_vec(), "png",
            ImageOptionsBuilder::new().x(3.2).y(1.0).w(2.5).h(1.5).rotate(30.0).build().0,
        );
        s.add_text("image rotate=30°", TextOptionsBuilder::new().x(3.2).y(2.55).w(2.5).h(0.4)
            .font_size(11.0).align(AlignH::Center).build());

        // Image: flip_h + flip_v
        s.add_image(
            RED_1PX_PNG.to_vec(), "png",
            ImageOptionsBuilder::new().x(6.2).y(1.0).w(2.5).h(1.5).flip_h().flip_v().build().0,
        );
        s.add_text("image flip_h + flip_v", TextOptionsBuilder::new().x(6.2).y(2.55).w(2.5).h(0.4)
            .font_size(11.0).align(AlignH::Center).build());

        // Text: normal
        s.add_text("Normal text box",
            TextOptionsBuilder::new().x(0.3).y(3.2).w(2.5).h(0.8)
                .font_size(15.0).fill("#4472C4").color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle).build());

        // Text: rotate 30°
        s.add_text("Rotated 30°",
            TextOptionsBuilder::new().x(3.2).y(3.2).w(2.5).h(0.8)
                .font_size(15.0).fill("#ED7D31").color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle)
                .rotate(30.0).build());

        // Text: flip_h
        s.add_text("Flipped H",
            TextOptionsBuilder::new().x(6.2).y(3.2).w(2.5).h(0.8)
                .font_size(15.0).fill("#70AD47").color("#FFFFFF")
                .align(AlignH::Center).valign(AlignV::Middle)
                .flip_h().build());
    }

    // ══════════════════════════════════════════════════════════
    // Milestone 1c/d/e — Text border, transparency, strike, underline color
    // ══════════════════════════════════════════════════════════
    {
        use deckmint::types::ShapeLineProps;
        let s = pres.add_slide();
        heading(s, "New: Text Border · Transparency · Strike · Underline Color");

        // Text box with border line
        s.add_text("Text box with a 2pt blue dashed border",
            TextOptionsBuilder::new()
                .x(0.4).y(0.95).w(5.8).h(0.6)
                .font_size(16.0)
                .line(ShapeLineProps {
                    color: Some("#4472C4".to_string()),
                    width: Some(2.0),
                    dash_type: Some("dash".to_string()),
                    ..Default::default()
                })
                .build());

        s.add_text("Border via .line_color() + .line_width() builder shorthand",
            TextOptionsBuilder::new()
                .x(0.4).y(1.65).w(5.8).h(0.6)
                .font_size(16.0)
                .line_color("#C00000").line_width(1.5)
                .build());

        // Text transparency
        let y = 2.45_f64;
        s.add_text_runs(
            vec![
                TextRunBuilder::new("Opaque  ").font_size(22.0).color("#222222").build(),
                TextRunBuilder::new("50% trans  ").font_size(22.0).color("#C45911").transparency(50.0).build(),
                TextRunBuilder::new("75% trans  ").font_size(22.0).color("#C45911").transparency(75.0).build(),
                TextRunBuilder::new("90% trans").font_size(22.0).color("#C45911").transparency(90.0).build(),
            ],
            TextOptionsBuilder::new().x(0.4).y(y).w(9.0).h(0.6).build(),
        );

        // Single vs double strike
        let y = 3.2_f64;
        s.add_text_runs(
            vec![
                TextRunBuilder::new("Single strike  ").font_size(22.0).strike().build(),
                TextRunBuilder::new("Double strike").font_size(22.0).strike_double().build(),
            ],
            TextOptionsBuilder::new().x(0.4).y(y).w(9.0).h(0.6).build(),
        );

        // Underline with custom color
        let y = 3.95_f64;
        s.add_text_runs(
            vec![
                TextRunBuilder::new("Red underline  ").font_size(22.0)
                    .underline("sng").underline_color("#FF0000").build(),
                TextRunBuilder::new("Green heavy  ").font_size(22.0)
                    .underline("heavy").underline_color("#00B050").build(),
                TextRunBuilder::new("Blue wavy").font_size(22.0)
                    .underline("wavy").underline_color("#4472C4").build(),
            ],
            TextOptionsBuilder::new().x(0.4).y(y).w(9.0).h(0.6).build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Milestone 2a — Arc / angle-range shapes
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "New: Arc / Angle-Range Shapes (PIE, ARC, BLOCK_ARC)");

        // Row 1: Pie slices at different sweep angles
        let sweeps: &[(f64, f64, &str, &str)] = &[
            (0.0,  90.0,  "#4472C4", "Pie 90°"),
            (0.0, 180.0,  "#ED7D31", "Pie 180°"),
            (0.0, 270.0,  "#70AD47", "Pie 270°"),
        ];
        for (i, (start, swing, color, label)) in sweeps.iter().enumerate() {
            let x = 0.4 + i as f64 * 3.1;
            s.add_shape(ShapeType::Pie,
                ShapeOptionsBuilder::new()
                    .x(x).y(1.0).w(2.5).h(2.5)
                    .fill_color(*color)
                    .angle_range(*start, *swing)
                    .build());
            s.add_text(*label,
                TextOptionsBuilder::new().x(x).y(3.6).w(2.5).h(0.4)
                    .font_size(12.0).align(AlignH::Center).build());
        }

        // Row 2: BlockArc with different thicknesses
        let arcs: &[(f64, f64, &str, &str)] = &[
            (0.0, 180.0, "#C45911", "BlockArc thick"),
            (0.0, 270.0, "#7030A0", "BlockArc thin"),
        ];
        for (i, (start, swing, color, label)) in arcs.iter().enumerate() {
            let x = 0.4 + i as f64 * 4.7;
            let thickness = if i == 0 { 0.6 } else { 0.15 };
            s.add_shape(ShapeType::BlockArc,
                ShapeOptionsBuilder::new()
                    .x(x).y(4.2).w(3.2).h(2.0)
                    .fill_color(*color)
                    .angle_range(*start, *swing)
                    .arc_thickness(thickness)
                    .build());
            s.add_text(*label,
                TextOptionsBuilder::new().x(x).y(6.25).w(3.2).h(0.4)
                    .font_size(12.0).align(AlignH::Center).build());
        }
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Pattern Fill
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Pattern Fill");

        let patterns = [
            (PatternType::Cross,     "Cross",     "#4472C4", "#FFFFFF"),
            (PatternType::DotGrid,   "DotGrid",   "#ED7D31", "#FFFFFF"),
            (PatternType::HorzBrick, "HorzBrick", "#70AD47", "#FFF2CC"),
            (PatternType::Trellis,   "Trellis",   "#9966CC", "#E2EFDA"),
            (PatternType::Wave,      "Wave",      "#C00000", "#FFFFFF"),
            (PatternType::ZigZag,    "ZigZag",    "#0070C0", "#DEEAF1"),
            (PatternType::LgConfetti,"LgConfetti","#FF0000", "#FFFF00"),
            (PatternType::Sphere,    "Sphere",    "#333333", "#EEEEEE"),
        ];

        for (i, &(ref pat, label, fg, bg)) in patterns.iter().enumerate() {
            let col = i % 4;
            let row = i / 4;
            let x = 0.3 + col as f64 * 2.45;
            let y = 1.0 + row as f64 * 1.6;

            s.add_shape(
                ShapeType::Rect,
                ShapeOptionsBuilder::new()
                    .x(x).y(y).w(2.2).h(1.2)
                    .pattern_fill(PatternFill {
                        pattern: pat.clone(),
                        fg_color: fg.to_string(),
                        bg_color: bg.to_string(),
                    })
                    .line_color("#CCCCCC").line_width(0.5)
                    .build(),
            );
            s.add_text(
                label,
                TextOptionsBuilder::new()
                    .x(x).y(y + 1.22).w(2.2).h(0.3)
                    .font_size(10.0).align(AlignH::Center).color("#333333")
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Milestone 2b — Custom shape geometry (freeform)
    // ══════════════════════════════════════════════════════════
    {
        use deckmint::CustomGeomPoint as CGP;
        let s = pres.add_slide();
        heading(s, "New: Custom Shape Geometry (Freeform Points)");

        // Triangle
        s.add_shape(ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .x(0.3).y(1.0).w(2.8).h(2.5)
                .fill_color("#4472C4")
                .custom_geometry(vec![
                    CGP::MoveTo(0.5, 0.0),
                    CGP::LineTo(1.0, 1.0),
                    CGP::LineTo(0.0, 1.0),
                    CGP::Close,
                ])
                .build());
        s.add_text("Triangle\n(3 LineTo points)", TextOptionsBuilder::new()
            .x(0.3).y(3.6).w(2.8).h(0.6).font_size(12.0).align(AlignH::Center).build());

        // Arrow pointing right (using LineTo polygon)
        s.add_shape(ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .x(3.5).y(1.0).w(2.8).h(2.5)
                .fill_color("#ED7D31")
                .custom_geometry(vec![
                    CGP::MoveTo(0.0, 0.25),
                    CGP::LineTo(0.6, 0.25),
                    CGP::LineTo(0.6, 0.0),
                    CGP::LineTo(1.0, 0.5),
                    CGP::LineTo(0.6, 1.0),
                    CGP::LineTo(0.6, 0.75),
                    CGP::LineTo(0.0, 0.75),
                    CGP::Close,
                ])
                .build());
        s.add_text("Arrow right\n(LineTo polygon)", TextOptionsBuilder::new()
            .x(3.5).y(3.6).w(2.8).h(0.6).font_size(12.0).align(AlignH::Center).build());

        // Rounded corner shape using CubicBezTo
        s.add_shape(ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .x(6.7).y(1.0).w(2.8).h(2.5)
                .fill_color("#70AD47")
                .custom_geometry(vec![
                    CGP::MoveTo(0.15, 0.0),
                    CGP::LineTo(0.85, 0.0),
                    CGP::CubicBezTo(1.0, 0.0, 1.0, 0.0, 1.0, 0.15),
                    CGP::LineTo(1.0, 0.85),
                    CGP::CubicBezTo(1.0, 1.0, 1.0, 1.0, 0.85, 1.0),
                    CGP::LineTo(0.15, 1.0),
                    CGP::CubicBezTo(0.0, 1.0, 0.0, 1.0, 0.0, 0.85),
                    CGP::LineTo(0.0, 0.15),
                    CGP::CubicBezTo(0.0, 0.0, 0.0, 0.0, 0.15, 0.0),
                    CGP::Close,
                ])
                .build());
        s.add_text("Rounded rect\n(CubicBezTo corners)", TextOptionsBuilder::new()
            .x(6.7).y(3.6).w(2.8).h(0.6).font_size(12.0).align(AlignH::Center).build());
    }

    // ══════════════════════════════════════════════════════════
    // Milestone 2c — Shape alt text
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "New: Shape Alt Text (Accessibility)");

        s.add_shape(ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .x(0.5).y(1.2).w(4.0).h(2.0)
                .fill_color("#4472C4")
                .alt_text("A blue rounded rectangle used as a decorative accent")
                .build());
        s.add_text("▲ Has alt text: \"A blue rounded rectangle…\"",
            TextOptionsBuilder::new().x(0.5).y(3.3).w(4.0).h(0.5)
                .font_size(11.0).color("#666666").build());

        s.add_shape(ShapeType::Ellipse,
            ShapeOptionsBuilder::new()
                .x(5.5).y(1.2).w(3.5).h(2.0)
                .fill_color("#ED7D31")
                .alt_text("Orange ellipse — chart placeholder")
                .build());
        s.add_text("▲ Has alt text: \"Orange ellipse — chart placeholder\"",
            TextOptionsBuilder::new().x(5.5).y(3.3).w(4.0).h(0.5)
                .font_size(11.0).color("#666666").build());

        s.add_text(
            "Alt text is set via .alt_text(\"...\") on ShapeOptionsBuilder.\n\
             It maps to descr=\"...\" on <p:cNvPr> and is read by screen readers.",
            TextOptionsBuilder::new().x(0.4).y(4.2).w(9.0).h(1.0)
                .font_size(14.0).color("#444444").build());
    }

    // ══════════════════════════════════════════════════════════
    // Milestone 1g — Slide background transparency
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color_transparency("#4472C4", 60.0);
        heading(s, "New: Background Color + Transparency");
        s.add_text(
            "Background: #4472C4 at 60% transparency\n(set via set_background_color_transparency())",
            TextOptionsBuilder::new()
                .x(0.5).y(2.0).w(9.0).h(1.5)
                .font_size(20.0).bold().align(AlignH::Center)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Theme Color Tints/Shades
    // ══════════════════════════════════════════════════════════
    {
        use deckmint::Color;

        let s = pres.add_slide();
        heading(s, "Theme Color Tints & Shades  (Color::ThemedWith)");

        let mods: &[(&str, ThemeColorMod)] = &[
            ("Accent1 (base)",          ThemeColorMod::default()),
            ("+ tint 40%",              ThemeColorMod::tint(40000)),
            ("+ tint 70%",              ThemeColorMod::tint(70000)),
            ("+ shade 25%",             ThemeColorMod::shade(25000)),
            ("+ shade 50%",             ThemeColorMod::shade(50000)),
            ("lum 75% / off 15%",       ThemeColorMod::lum(75000, 15000)),
        ];

        for (i, &(label, ref m)) in mods.iter().enumerate() {
            let x = 0.3 + (i % 3) as f64 * 3.2;
            let y = 1.0 + (i / 3) as f64 * 1.6;

            let fill_color = Color::ThemedWith(SchemeColor::Accent1, m.clone());
            s.add_shape(
                ShapeType::Rect,
                ShapeOptionsBuilder::new()
                    .x(x).y(y).w(2.9).h(1.0)
                    .fill_color_value(fill_color)
                    .line_color("#CCCCCC").line_width(0.5)
                    .build(),
            );
            s.add_text(
                label,
                TextOptionsBuilder::new()
                    .x(x).y(y + 1.02).w(2.9).h(0.4)
                    .font_size(10.5).align(AlignH::Center).color("#333333")
                    .build(),
            );
        }
        s.add_text(
            "Color::ThemedWith(SchemeColor::Accent1, ThemeColorMod::tint(40000))\n\
             emits <a:schemeClr val=\"accent1\"><a:tint val=\"40000\"/></a:schemeClr>",
            TextOptionsBuilder::new()
                .x(0.3).y(4.5).w(9.4).h(0.9)
                .font_size(11.0).color("#666666").italic()
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Milestone 3a — Slide numbers
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "New: Slide Numbers");
        s.set_slide_number(SlideNumberProps {
            x: Some(deckmint::Coord::Inches(8.5)),
            y: Some(deckmint::Coord::Inches(7.0)),
            w: Some(deckmint::Coord::Inches(1.5)),
            h: Some(deckmint::Coord::Inches(0.4)),
            font_size: Some(14.0),
            color: Some("#444444".to_string()),
            bold: true,
            align: Some("r".to_string()),
            ..Default::default()
        });
        s.add_text(
            "This slide has a slide-number placeholder in the bottom-right corner.\n\
             Set via slide.set_slide_number(SlideNumberProps { x, y, w, h, font_size, color, … })\n\
             The number is rendered automatically by PowerPoint/LibreOffice.",
            TextOptionsBuilder::new()
                .x(0.5).y(1.5).w(9.0).h(2.5)
                .font_size(16.0)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Advanced Hyperlink Actions
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Advanced Hyperlink Actions");

        let actions: &[(&str, HyperlinkAction, &str)] = &[
            ("Next Slide →",   HyperlinkAction::NextSlide,  "#4472C4"),
            ("← Prev Slide",   HyperlinkAction::PrevSlide,  "#ED7D31"),
            ("⏮ First Slide",  HyperlinkAction::FirstSlide, "#70AD47"),
            ("⏭ Last Slide",   HyperlinkAction::LastSlide,  "#9966CC"),
            ("✕ End Show",     HyperlinkAction::EndShow,    "#C00000"),
        ];

        for (i, &(label, ref action, color)) in actions.iter().enumerate() {
            let x = 0.5 + (i % 3) as f64 * 3.1;
            let y = 1.1 + (i / 3) as f64 * 1.5;
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .x(x).y(y).w(2.7).h(0.9)
                    .fill_color(color)
                    .hyperlink(HyperlinkProps {
                        r_id: 0,
                        slide: None,
                        url: None,
                        tooltip: Some(format!("Action: {label}")),
                        action: Some(action.clone()),
                    })
                    .build(),
            );
            s.add_text(
                label,
                TextOptionsBuilder::new()
                    .x(x).y(y + 0.15).w(2.7).h(0.6)
                    .font_size(14.0).bold().align(AlignH::Center).color("#FFFFFF")
                    .build(),
            );
        }
        s.add_text(
            "Click any button in Slideshow mode — each fires a navigation action with no URL relationship.",
            TextOptionsBuilder::new()
                .x(0.3).y(4.3).w(9.4).h(0.5)
                .font_size(11.0).color("#666666").italic()
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Slide Transitions
    // ══════════════════════════════════════════════════════════
    {
        // Apply a fade transition to this slide
        let s = pres.add_slide();
        s.set_transition(TransitionProps {
            transition_type: deckmint::TransitionType::Fade,
            speed: Some(deckmint::TransitionSpeed::Medium),
            advance_on_click: true,
            ..Default::default()
        });
        heading(s, "Slide Transitions  (this slide: Fade)");
        s.add_text(
            "This slide has a Fade transition applied via set_transition().\n\
             Switch to Slideshow mode and advance to see it.",
            TextOptionsBuilder::new()
                .x(0.5).y(1.1).w(9.0).h(1.2).font_size(18.0).build(),
        );

        let types_text = "Supported types: Cut, Fade, Push, Wipe, Split, Cover, Uncover,\n\
             Zoom, Flash, Morph, Vortex, Ripple, Glitter, Honeycomb, Shred, Switch,\n\
             Flip, Pan, Ferris, Gallery, Conveyor, Doors, Box, Random, RandomBar,\n\
             Circle, Diamond, Wheel, Checker, Blinds, Strips, Plus (30 types total)";
        s.add_text(
            types_text,
            TextOptionsBuilder::new()
                .x(0.5).y(2.5).w(9.0).h(2.0).font_size(14.0).color("#444444").build(),
        );
    }

    // ── Next slide uses Push transition ────────────────────
    {
        let s = pres.add_slide();
        s.set_transition(TransitionProps::push(TransitionDir::Left));
        heading(s, "Slide Transitions  (this slide: Push Left)");
        s.add_text(
            "TransitionProps::push(TransitionDir::Left)\n\n\
             Direction variants: Left, Right, Up, Down, LeftDown, LeftUp, RightDown, RightUp\n\
             Speed variants: Slow, Medium, Fast\n\
             Auto-advance: set advance_after_ms to advance without a click",
            TextOptionsBuilder::new()
                .x(0.5).y(1.1).w(9.0).h(3.0).font_size(16.0).build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Connectors
    // ══════════════════════════════════════════════════════════
    {
        use deckmint::ShapeLineProps;

        let s = pres.add_slide();
        heading(s, "Connectors  (<p:cxnSp>)");

        // Three boxes to connect
        let boxes = [(1.0_f64, 1.2_f64), (4.5_f64, 1.2_f64), (8.0_f64, 1.2_f64)];
        let labels = ["Box A", "Box B", "Box C"];
        for ((bx, by), &lbl) in boxes.iter().zip(labels.iter()) {
            s.add_shape(
                ShapeType::Rect,
                ShapeOptionsBuilder::new()
                    .x(*bx).y(*by).w(1.5).h(0.8)
                    .fill_color("#4472C4").line_color("#1F3864").line_width(1.5)
                    .build(),
            );
            s.add_text(
                lbl,
                TextOptionsBuilder::new()
                    .x(*bx).y(*by + 0.1).w(1.5).h(0.6)
                    .font_size(14.0).bold().align(AlignH::Center).color("#FFFFFF")
                    .build(),
            );
        }

        // Straight connector: A → B
        s.add_connector(
            ConnectorType::Straight,
            ConnectorOptions {
                x1: Some(deckmint::Coord::Inches(2.5)),
                y1: Some(deckmint::Coord::Inches(1.6)),
                x2: Some(deckmint::Coord::Inches(4.5)),
                y2: Some(deckmint::Coord::Inches(1.6)),
                line: Some(ShapeLineProps {
                    color: Some("#ED7D31".to_string()),
                    width: Some(2.0),
                    end_arrow_type: Some("triangle".to_string()),
                    ..Default::default()
                }),
                ..Default::default()
            },
        );

        // Elbow connector: B → C
        s.add_connector(
            ConnectorType::Elbow,
            ConnectorOptions {
                x1: Some(deckmint::Coord::Inches(6.0)),
                y1: Some(deckmint::Coord::Inches(1.6)),
                x2: Some(deckmint::Coord::Inches(8.0)),
                y2: Some(deckmint::Coord::Inches(1.6)),
                line: Some(ShapeLineProps {
                    color: Some("#70AD47".to_string()),
                    width: Some(2.0),
                    end_arrow_type: Some("triangle".to_string()),
                    ..Default::default()
                }),
                ..Default::default()
            },
        );

        // Curved connector: A → C (diagonal)
        s.add_connector(
            ConnectorType::Curved,
            ConnectorOptions {
                x1: Some(deckmint::Coord::Inches(1.75)),
                y1: Some(deckmint::Coord::Inches(2.0)),
                x2: Some(deckmint::Coord::Inches(8.75)),
                y2: Some(deckmint::Coord::Inches(3.5)),
                line: Some(ShapeLineProps {
                    color: Some("#9966CC".to_string()),
                    width: Some(2.0),
                    dash_type: Some("dash".to_string()),
                    end_arrow_type: Some("triangle".to_string()),
                    ..Default::default()
                }),
                ..Default::default()
            },
        );

        s.add_text(
            "Orange: Straight  |  Green: Elbow  |  Purple dashed: Curved\n\
             add_connector(ConnectorType::Straight/Elbow/Curved, ConnectorOptions { x1, y1, x2, y2, line, … })",
            TextOptionsBuilder::new()
                .x(0.3).y(4.1).w(9.4).h(0.9)
                .font_size(11.0).color("#666666").italic()
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Milestone 3b — Sections
    // ══════════════════════════════════════════════════════════
    {
        pres.add_section("Section A — Introduction");
        let s = pres.add_slide();
        heading(s, "Sections: Slide in Section A");
        s.add_text(
            "This slide belongs to \"Section A — Introduction\".\n\
             Sections are added via pres.add_section(name) before the relevant add_slide() calls.\n\
             They appear as collapsible groups in PowerPoint's slide panel.",
            TextOptionsBuilder::new()
                .x(0.5).y(1.2).w(9.0).h(2.5)
                .font_size(16.0)
                .build(),
        );

        pres.add_section("Section B — Content");
        let s = pres.add_slide();
        heading(s, "Sections: First Slide in Section B");
        s.add_text(
            "This slide belongs to \"Section B — Content\".",
            TextOptionsBuilder::new()
                .x(0.5).y(1.5).w(9.0).h(1.0)
                .font_size(18.0)
                .build(),
        );

        let s = pres.add_slide();
        heading(s, "Sections: Second Slide in Section B");
        s.add_text(
            "Also in Section B. Multiple slides can share a section.",
            TextOptionsBuilder::new()
                .x(0.5).y(1.5).w(9.0).h(1.0)
                .font_size(18.0)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Table Style Preset + AlignH::Distribute
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Table Style Preset + AlignH::Distribute");

        let rows = vec![
            vec![
                TableCell::new("Quarter"),
                TableCell::new("Revenue"),
                TableCell::new("Growth"),
            ],
            vec![TableCell::new("Q1"), TableCell::new("$1.2M"), TableCell::new("+8%")],
            vec![TableCell::new("Q2"), TableCell::new("$1.5M"), TableCell::new("+25%")],
            vec![TableCell::new("Q3"), TableCell::new("$1.4M"), TableCell::new("−7%")],
            vec![TableCell::new("Q4"), TableCell::new("$1.9M"), TableCell::new("+36%")],
        ];
        s.add_table(
            rows,
            TableOptionsBuilder::new()
                .x(1.0).y(1.0).w(8.0).h(3.0)
                // Medium Style 2 - Accent 1
                .table_style_id("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
                .align(AlignH::Center)
                .build(),
        );

        s.add_text(
            "table_style_id(\"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}\")  →  OOXML built-in style applied",
            TextOptionsBuilder::new()
                .x(0.3).y(4.1).w(9.4).h(0.4)
                .font_size(11.0).color("#666666").italic()
                .build(),
        );

        // AlignH::Distribute demo
        s.add_text(
            "AlignH::Distribute  →  text distributed evenly across width",
            TextOptionsBuilder::new()
                .x(0.5).y(4.6).w(9.0).h(0.5)
                .font_size(14.0).align(AlignH::Distribute).color("#2E4057")
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Animation Sequencing (WithPrevious / AfterPrevious)
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Animation Sequencing  (WithPrevious / AfterPrevious)");

        // Shape 1: appears on click
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .x(1.0).y(1.5).w(2.5).h(1.5)
                .fill_color("#4472C4")
                .build(),
        );
        s.add_text(
            "1 — On Click",
            TextOptionsBuilder::new()
                .x(1.0).y(1.5).w(2.5).h(1.5)
                .font_size(14.0).color("#FFFFFF").align(AlignH::Center).valign(AlignV::Middle)
                .animation(AnimationEffect::fade_in())
                .build(),
        );
        // Shape 2: fades in WITH shape 1 (no extra click)
        s.add_text(
            "2 — With Previous",
            TextOptionsBuilder::new()
                .x(4.0).y(1.5).w(2.5).h(1.5)
                .font_size(14.0).color("#FFFFFF").align(AlignH::Center).valign(AlignV::Middle)
                .fill("#ED7D31".to_string())
                .animation(AnimationEffect::fade_in().with_previous())
                .build(),
        );
        // Shape 3: flies in AFTER shapes 1+2 finish (no click, auto)
        s.add_text(
            "3 — After Previous",
            TextOptionsBuilder::new()
                .x(7.0).y(1.5).w(2.5).h(1.5)
                .font_size(14.0).color("#FFFFFF").align(AlignH::Center).valign(AlignV::Middle)
                .fill("#70AD47".to_string())
                .animation(AnimationEffect::fly_in(Direction::Down).after_previous())
                .build(),
        );
        // Shape 4: after previous + 500ms delay
        s.add_text(
            "4 — After + 500ms delay",
            TextOptionsBuilder::new()
                .x(4.0).y(3.5).w(2.5).h(1.5)
                .font_size(14.0).color("#FFFFFF").align(AlignH::Center).valign(AlignV::Middle)
                .fill("#A855C7".to_string())
                .animation(AnimationEffect::zoom_in().after_previous().delay(500))
                .build(),
        );

        s.add_text(
            "Run slideshow (F5) → single click triggers shapes 1+2 together,\n\
             then 3 auto-plays after they finish, then 4 after a 500 ms delay.",
            TextOptionsBuilder::new()
                .x(0.3).y(5.1).w(9.4).h(0.7)
                .font_size(11.0).color("#666666").italic()
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Field Codes (Slide Number / Date)
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Field Codes  (FieldType::SlideNumber / DateTime)");

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Slide number: ").font_size(20.0).build(),
                TextRunBuilder::new("").font_size(20.0).bold().field(FieldType::SlideNumber).build(),
            ],
            TextOptionsBuilder::new().x(1.0).y(1.5).w(8.0).h(1.0).build(),
        );

        s.add_text_runs(
            vec![
                TextRunBuilder::new("Date/time: ").font_size(20.0).build(),
                TextRunBuilder::new("").font_size(20.0).bold().field(FieldType::DateTime).build(),
            ],
            TextOptionsBuilder::new().x(1.0).y(2.8).w(8.0).h(1.0).build(),
        );

        s.add_text(
            "The fields above auto-update when opened in PowerPoint.\n\
             TextRunBuilder::new(\"\").field(FieldType::SlideNumber).build()",
            TextOptionsBuilder::new()
                .x(0.3).y(4.5).w(9.4).h(0.9)
                .font_size(11.0).color("#666666").italic()
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Multi-Column Text
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Multiple Text Columns  (.columns(N))");

        let lorem = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. \
            Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. \
            Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris \
            nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in \
            reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla \
            pariatur. Excepteur sint occaecat cupidatat non proident.";
        s.add_text(
            lorem,
            TextOptionsBuilder::new()
                .x(0.5).y(1.2).w(9.0).h(2.5)
                .font_size(14.0).color("#333333")
                .columns(2).column_spacing(0.4)
                .build(),
        );
        s.add_text(
            "Three columns:",
            TextOptionsBuilder::new()
                .x(0.5).y(3.8).w(9.0).h(0.4)
                .font_size(12.0).bold().color("#2E4057")
                .build(),
        );
        s.add_text(
            lorem,
            TextOptionsBuilder::new()
                .x(0.5).y(4.2).w(9.0).h(1.8)
                .font_size(11.0).color("#333333")
                .columns(3).column_spacing(0.3)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Gradient Stops with Theme Colors
    // ══════════════════════════════════════════════════════════
    {
        use deckmint::Color;

        let s = pres.add_slide();
        heading(s, "Gradient Stops with Theme Colors  (GradientStop::from_color)");

        // Gradient using theme colours with tint/shade modifiers
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .x(0.5).y(1.5).w(9.0).h(2.0)
                .gradient_fill(GradientFill::linear(0.0, vec![
                    GradientStop::from_color(Color::ThemedWith(SchemeColor::Accent1, ThemeColorMod::tint(40000)), 0.0),
                    GradientStop::from_color(Color::Theme(SchemeColor::Accent1), 50.0),
                    GradientStop::from_color(Color::ThemedWith(SchemeColor::Accent1, ThemeColorMod::shade(50000)), 100.0),
                ]))
                .build(),
        );
        s.add_text(
            "Accent1 tint 40% → Accent1 → Accent1 shade 50%",
            TextOptionsBuilder::new()
                .x(0.5).y(3.6).w(9.0).h(0.5)
                .font_size(12.0).color("#333333").align(AlignH::Center)
                .build(),
        );

        // Mixed hex + theme gradient
        s.add_shape(
            ShapeType::Rect,
            ShapeOptionsBuilder::new()
                .x(0.5).y(4.2).w(9.0).h(1.2)
                .gradient_fill(GradientFill::linear(90.0, vec![
                    GradientStop::new("#FFFFFF", 0.0),
                    GradientStop::from_color(Color::Theme(SchemeColor::Accent2), 100.0),
                ]))
                .build(),
        );
        s.add_text(
            "White (hex) → Accent2 (theme) at 90°",
            TextOptionsBuilder::new()
                .x(0.5).y(5.5).w(9.0).h(0.4)
                .font_size(12.0).color("#333333").align(AlignH::Center)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // NEW — Image Color Adjustments
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        heading(s, "Image Color Adjustments  (.brightness / .contrast / .grayscale)");

        s.add_text(
            "ImageOptionsBuilder::new()\n  .brightness(30.0)    // brighter\n  \
             .contrast(-20.0)     // less contrast\n  .grayscale()         // b&w",
            TextOptionsBuilder::new()
                .x(3.0).y(1.5).w(6.5).h(2.5)
                .font_size(14.0).color("#2E4057")
                .build(),
        );

        s.add_text(
            "Emits <a:lum bright=\"30000\" contrast=\"-20000\"/> and <a:grayscl/> inside <a:blip>",
            TextOptionsBuilder::new()
                .x(0.3).y(4.5).w(9.4).h(0.9)
                .font_size(11.0).color("#666666").italic()
                .build(),
        );
    }

    // ── Milestone 4: Charts ──────────────────────────────────
    {
        let labels = vec!["Q1", "Q2", "Q3", "Q4"];
        let series = vec![
            ChartSeries::new("North", labels.clone(), vec![4.3, 2.5, 3.5, 4.5]),
            ChartSeries::new("South", labels.clone(), vec![2.4, 4.4, 1.8, 2.8]),
            ChartSeries::new("East",  labels.clone(), vec![2.0, 2.0, 3.0, 5.0]),
        ];

        // Bar / Column chart
        let s = pres.add_slide();
        heading(s, "Charts: Column Chart");
        s.add_chart(
            ChartType::Bar,
            series.clone(),
            ChartOptionsBuilder::new()
                .x(0.5).y(1.2).w(9.0).h(4.3)
                .title("Quarterly Sales (Column)")
                .show_value()
                .build(),
        );

        // Horizontal bar chart
        let s = pres.add_slide();
        heading(s, "Charts: Horizontal Bar Chart");
        s.add_chart(
            ChartType::Bar,
            series.clone(),
            ChartOptionsBuilder::new()
                .x(0.5).y(1.2).w(9.0).h(4.3)
                .title("Quarterly Sales (Bar)")
                .bar_dir(BarDir::Bar)
                .bar_grouping(BarGrouping::Stacked)
                .build(),
        );

        // Line chart
        let s = pres.add_slide();
        heading(s, "Charts: Line Chart");
        s.add_chart(
            ChartType::Line,
            series.clone(),
            ChartOptionsBuilder::new()
                .x(0.5).y(1.2).w(9.0).h(4.3)
                .title("Quarterly Trend (Line)")
                .line_smooth()
                .cat_axis_title("Quarter")
                .val_axis_title("Units (M)")
                .build(),
        );

        // Pie chart
        let pie_labels = vec!["Apples", "Bananas", "Cherries", "Dates"];
        let pie_series = vec![ChartSeries::new("Fruit", pie_labels, vec![28.0, 17.0, 40.0, 15.0])];
        let s = pres.add_slide();
        heading(s, "Charts: Pie Chart");
        s.add_chart(
            ChartType::Pie,
            pie_series.clone(),
            ChartOptionsBuilder::new()
                .x(1.0).y(1.2).w(8.0).h(4.3)
                .title("Fruit Distribution")
                .show_value()
                .first_slice_angle(45)
                .build(),
        );

        // Doughnut chart
        let s = pres.add_slide();
        heading(s, "Charts: Doughnut Chart");
        s.add_chart(
            ChartType::Doughnut,
            pie_series,
            ChartOptionsBuilder::new()
                .x(1.0).y(1.2).w(8.0).h(4.3)
                .title("Fruit Distribution (Doughnut)")
                .hole_size(60)
                .legend_pos(LegendPos::Right)
                .build(),
        );

        // Area chart
        let s = pres.add_slide();
        heading(s, "Charts: Area Chart");
        s.add_chart(
            ChartType::Area,
            series.clone(),
            ChartOptionsBuilder::new()
                .x(0.5).y(1.2).w(9.0).h(4.3)
                .title("Quarterly Sales (Area)")
                .val_axis_min(0.0)
                .val_axis_max(6.0)
                .build(),
        );

        // Scatter chart
        let scatter_series = vec![
            ChartSeries::new("Dataset A", Vec::<String>::new(), vec![1.0, 4.0, 9.0, 16.0, 25.0]),
            ChartSeries::new("Dataset B", Vec::<String>::new(), vec![2.0, 3.0, 5.0, 8.0, 13.0]),
        ];
        let s = pres.add_slide();
        heading(s, "Charts: Scatter Chart");
        s.add_chart(
            ChartType::Scatter,
            scatter_series,
            ChartOptionsBuilder::new()
                .x(0.5).y(1.2).w(9.0).h(4.3)
                .title("Scatter Plot")
                .no_grid_lines()
                .build(),
        );
    }

    let out = "features.pptx";
    pres.write_to_file(out).expect("write failed");
    println!("Wrote {out}");
    println!("Open in LibreOffice Impress / PowerPoint to review all features.");
}
