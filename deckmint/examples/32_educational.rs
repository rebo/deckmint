//! 5-slide solar system lesson — dark space theme with bright planet-colored
//! accents, callout shapes, data table, pie chart, and quiz questions.

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::table::{TableCell, TableOptionsBuilder};
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{
    AlignH, AlignV, ChartOptionsBuilder, ChartSeries, ChartType, Presentation, ShapeType,
};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "The Solar System".to_string();

    // ── Colors ────────────────────────────────────────────────
    let space_bg = "#0B0C10";
    let dark_card = "#1A1B2E";
    let white = "#FFFFFF";
    let muted = "#8892A0";
    let sun_color = "#FFD700";
    let mercury_color = "#B0B0B0";
    let venus_color = "#E8A84C";
    let earth_color = "#4488CC";
    let mars_color = "#CC4422";
    let jupiter_color = "#D4A574";
    let saturn_color = "#C4A35A";
    let uranus_color = "#7BC8C8";
    let neptune_color = "#4466AA";

    // ══════════════════════════════════════════════════════════
    // Slide 1: Title with planet circles
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color(space_bg);

        // Sun (partial, left side)
        s.add_shape(
            ShapeType::Ellipse,
            ShapeOptionsBuilder::new()
                .bounds(-0.5, 1.5, 2.5, 2.5)
                .fill_color(sun_color)
                .build(),
        );

        s.add_text(
            "THE SOLAR SYSTEM",
            TextOptionsBuilder::new()
                .bounds(2.0, 1.2, 7.5, 1.2)
                .font_size(44.0)
                .bold()
                .color(white)
                .align(AlignH::Center)
                .valign(AlignV::Bottom)
                .build(),
        );

        s.add_text(
            "A Journey Through Our Cosmic Neighborhood",
            TextOptionsBuilder::new()
                .bounds(2.0, 2.5, 7.5, 0.6)
                .font_size(18.0)
                .color(muted)
                .align(AlignH::Center)
                .valign(AlignV::Top)
                .build(),
        );

        // Row of planet circles across the bottom
        let planets = [
            (mercury_color, 0.2),
            (venus_color, 0.3),
            (earth_color, 0.32),
            (mars_color, 0.25),
            (jupiter_color, 0.6),
            (saturn_color, 0.5),
            (uranus_color, 0.4),
            (neptune_color, 0.38),
        ];

        let mut px = 1.2;
        let base_y = 3.8;
        for (color, size) in &planets {
            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(px, base_y + (0.6 - size) / 2.0, *size, *size)
                    .fill_color(*color)
                    .build(),
            );
            px += size + 0.8;
        }

        // Decorative stars (small ellipses)
        let stars = [
            (3.2, 0.8, 0.06),
            (7.5, 0.5, 0.04),
            (8.8, 1.5, 0.05),
            (5.0, 0.3, 0.03),
            (9.2, 3.2, 0.04),
            (1.5, 0.4, 0.05),
            (6.8, 4.8, 0.04),
        ];
        for (sx, sy, ss) in &stars {
            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(*sx, *sy, *ss, *ss)
                    .fill_color(white)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Planet data table
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color(space_bg);

        s.add_text(
            "PLANET DATA",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(sun_color)
                .build(),
        );

        s.add_text(
            "Key Facts About Our Planets",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.45, 9.0, 0.5)
                .font_size(26.0)
                .bold()
                .color(white)
                .build(),
        );

        let hdr_bg = "#1C2340";

        let header = vec![
            TableCell::new("Planet").bold().fill(hdr_bg).color(sun_color).align(AlignH::Center),
            TableCell::new("Diameter (km)").bold().fill(hdr_bg).color(sun_color).align(AlignH::Center),
            TableCell::new("Dist. from Sun (AU)").bold().fill(hdr_bg).color(sun_color).align(AlignH::Center),
            TableCell::new("Type").bold().fill(hdr_bg).color(sun_color).align(AlignH::Center),
        ];

        let planet_data = [
            ("Mercury", "4,879", "0.39", "Rocky", mercury_color),
            ("Venus", "12,104", "0.72", "Rocky", venus_color),
            ("Earth", "12,742", "1.00", "Rocky", earth_color),
            ("Mars", "6,779", "1.52", "Rocky", mars_color),
            ("Jupiter", "139,820", "5.20", "Gas Giant", jupiter_color),
            ("Saturn", "116,460", "9.58", "Gas Giant", saturn_color),
            ("Uranus", "50,724", "19.22", "Ice Giant", uranus_color),
            ("Neptune", "49,244", "30.07", "Ice Giant", neptune_color),
        ];

        let dark_row = "#0D0E18";
        let alt_row = "#141524";

        let mut rows = vec![header];
        for (i, (name, diameter, dist, ptype, color)) in planet_data.iter().enumerate() {
            let bg = if i % 2 == 0 { dark_row } else { alt_row };
            rows.push(vec![
                TableCell::new(*name).bold().fill(bg).color(*color),
                TableCell::new(*diameter).fill(bg).color(white).align(AlignH::Right),
                TableCell::new(*dist).fill(bg).color(white).align(AlignH::Center),
                TableCell::new(*ptype).fill(bg).color(muted).align(AlignH::Center),
            ]);
        }

        s.add_table(
            rows,
            TableOptionsBuilder::new()
                .bounds(0.5, 1.1, 9.0, 4.2)
                .col_w(vec![2.0, 2.3, 2.7, 2.0])
                .font_size(11.0)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Fun facts in callout shapes
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color(space_bg);

        s.add_text(
            "FUN FACTS",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(sun_color)
                .build(),
        );

        s.add_text(
            "Amazing Things About Space",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.45, 9.0, 0.5)
                .font_size(26.0)
                .bold()
                .color(white)
                .build(),
        );

        // Fact 1 — WedgeRoundRectCallout
        s.add_shape(
            ShapeType::WedgeRoundRectCallout,
            ShapeOptionsBuilder::new()
                .bounds(0.5, 1.2, 4.2, 1.6)
                .fill_color("#1A2744")
                .line_color(jupiter_color)
                .line_width(2.0)
                .build(),
        );
        s.add_text(
            "Jupiter is so big that all\nother planets could fit\ninside it!",
            TextOptionsBuilder::new()
                .bounds(0.8, 1.3, 3.6, 1.3)
                .font_size(13.0)
                .bold()
                .color(jupiter_color)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Planet icon for fact 1
        s.add_shape(
            ShapeType::Ellipse,
            ShapeOptionsBuilder::new()
                .bounds(2.2, 2.9, 0.5, 0.5)
                .fill_color(jupiter_color)
                .build(),
        );

        // Fact 2 — CloudCallout
        s.add_shape(
            ShapeType::CloudCallout,
            ShapeOptionsBuilder::new()
                .bounds(5.3, 1.1, 4.2, 1.8)
                .fill_color("#1A2744")
                .line_color(saturn_color)
                .line_width(2.0)
                .build(),
        );
        s.add_text(
            "Saturn's rings are made\nof ice and rock — some\npieces are as big as houses!",
            TextOptionsBuilder::new()
                .bounds(5.6, 1.3, 3.6, 1.3)
                .font_size(12.0)
                .bold()
                .color(saturn_color)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Fact 3 — WedgeEllipseCallout
        s.add_shape(
            ShapeType::WedgeEllipseCallout,
            ShapeOptionsBuilder::new()
                .bounds(0.8, 3.3, 4.0, 1.8)
                .fill_color("#1A2744")
                .line_color(mars_color)
                .line_width(2.0)
                .build(),
        );
        s.add_text(
            "Mars has the tallest\nvolcano in the solar\nsystem: Olympus Mons!",
            TextOptionsBuilder::new()
                .bounds(1.2, 3.5, 3.2, 1.2)
                .font_size(13.0)
                .bold()
                .color(mars_color)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Fact 4 — another WedgeRoundRectCallout
        s.add_shape(
            ShapeType::WedgeRoundRectCallout,
            ShapeOptionsBuilder::new()
                .bounds(5.5, 3.2, 4.0, 1.8)
                .fill_color("#1A2744")
                .line_color(earth_color)
                .line_width(2.0)
                .build(),
        );
        s.add_text(
            "A day on Venus is longer\nthan its year! It takes 243\nEarth days to rotate once.",
            TextOptionsBuilder::new()
                .bounds(5.8, 3.4, 3.4, 1.2)
                .font_size(12.0)
                .bold()
                .color(earth_color)
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 4: Pie chart of planet mass distribution
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color(space_bg);

        s.add_text(
            "MASS DISTRIBUTION",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(sun_color)
                .build(),
        );

        s.add_text(
            "Where Is All the Mass?",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.45, 9.0, 0.5)
                .font_size(26.0)
                .bold()
                .color(white)
                .build(),
        );

        // Pie chart — planet mass shares
        let pie_series = vec![ChartSeries::new(
            "Mass Share",
            vec!["Jupiter", "Saturn", "Neptune", "Uranus", "Earth", "Venus", "Mars", "Mercury"],
            vec![71.2, 21.3, 3.8, 2.6, 0.45, 0.37, 0.048, 0.025],
        )];

        s.add_chart(
            ChartType::Pie,
            pie_series,
            ChartOptionsBuilder::new()
                .bounds(0.5, 1.1, 5.5, 4.2)
                .title("Planetary Mass Distribution (%)")
                .show_value()
                .chart_colors(vec![
                    jupiter_color,
                    saturn_color,
                    neptune_color,
                    uranus_color,
                    earth_color,
                    venus_color,
                    mars_color,
                    mercury_color,
                ])
                .build(),
        );

        // Side panel with context
        s.add_shape(
            ShapeType::RoundRect,
            ShapeOptionsBuilder::new()
                .bounds(6.3, 1.2, 3.3, 3.8)
                .fill_color(dark_card)
                .rect_radius(0.08)
                .build(),
        );

        s.add_text(
            "Key Insight",
            TextOptionsBuilder::new()
                .bounds(6.5, 1.4, 2.9, 0.4)
                .font_size(16.0)
                .bold()
                .color(sun_color)
                .align(AlignH::Center)
                .build(),
        );

        s.add_text(
            "Jupiter alone contains more than twice the mass of all other planets combined!\n\nThe four inner rocky planets (Mercury, Venus, Earth, Mars) make up less than 1% of the total planetary mass.",
            TextOptionsBuilder::new()
                .bounds(6.5, 1.9, 2.9, 2.8)
                .font_size(11.0)
                .color(muted)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 5: Quiz — 3 numbered questions
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color(space_bg);

        s.add_text(
            "QUIZ TIME!",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.4)
                .font_size(12.0)
                .bold()
                .color(sun_color)
                .build(),
        );

        s.add_text(
            "Test Your Knowledge",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.45, 9.0, 0.5)
                .font_size(26.0)
                .bold()
                .color(white)
                .build(),
        );

        let questions = [
            (
                "Which planet is the largest in our solar system?",
                "A) Saturn    B) Jupiter    C) Neptune    D) Uranus",
                jupiter_color,
            ),
            (
                "How many planets in our solar system are classified as 'rocky'?",
                "A) 2    B) 3    C) 4    D) 5",
                earth_color,
            ),
            (
                "Which planet has the tallest volcano, Olympus Mons?",
                "A) Venus    B) Earth    C) Mars    D) Mercury",
                mars_color,
            ),
        ];

        let q_area = deckmint::layout::CellRect {
            x: 0.5,
            y: 1.2,
            w: 9.0,
            h: 3.9,
        };
        let q_rows = deckmint::layout::split_v(&q_area, 3, 0.2);

        for (i, (question, options, color)) in questions.iter().enumerate() {
            let row = q_rows[i];

            // Question card
            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(row)
                    .fill_color(dark_card)
                    .line_color(*color)
                    .line_width(1.5)
                    .rect_radius(0.08)
                    .build(),
            );

            // Number badge
            let badge_size = 0.45;
            let badge_x = row.x + 0.2;
            let badge_y = row.y + (row.h - badge_size) / 2.0;
            s.add_shape(
                ShapeType::Ellipse,
                ShapeOptionsBuilder::new()
                    .bounds(badge_x, badge_y, badge_size, badge_size)
                    .fill_color(*color)
                    .build(),
            );
            s.add_text(
                &format!("{}", i + 1),
                TextOptionsBuilder::new()
                    .bounds(badge_x, badge_y, badge_size, badge_size)
                    .font_size(16.0)
                    .bold()
                    .color(space_bg)
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Question text
            s.add_text(
                *question,
                TextOptionsBuilder::new()
                    .bounds(row.x + 0.85, row.y + 0.1, row.w - 1.1, 0.55)
                    .font_size(14.0)
                    .bold()
                    .color(white)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Answer options
            s.add_text(
                *options,
                TextOptionsBuilder::new()
                    .bounds(row.x + 0.85, row.y + 0.6, row.w - 1.1, 0.45)
                    .font_size(11.0)
                    .color(muted)
                    .valign(AlignV::Top)
                    .build(),
            );
        }
    }

    pres.write_to_file("32_educational.pptx").unwrap();
    println!("Wrote 32_educational.pptx");
}
