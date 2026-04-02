//! Flowchart with 3 connector types and a hub-and-spoke network diagram.

use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::types::ShapeLineProps;
use deckmint::{AlignH, AlignV, ConnectorOptionsBuilder, ConnectorType, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();

    // ══════════════════════════════════════════════════════════
    // Slide 1: Flowchart — 4 process steps with 3 connector types
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#F2F2F2");

        slide.add_text(
            "Flowchart: Connector Types",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.2, 9.0, 0.6)
                .font_size(24.0)
                .bold()
                .color("#333333")
                .build(),
        );

        // Step boxes — rounded rectangles in a Z-pattern
        let steps = [
            ("1. Plan",    1.0,  1.3, "#4472C4"),
            ("2. Design",  5.0,  1.3, "#ED7D31"),
            ("3. Build",   1.0,  3.5, "#70AD47"),
            ("4. Deploy",  5.0,  3.5, "#FFC000"),
        ];
        let box_w = 3.0;
        let box_h = 1.2;

        for (label, x, y, color) in &steps {
            slide.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(*x, *y, box_w, box_h)
                    .fill_color(*color)
                    .rect_radius(0.15)
                    .build(),
            );
            slide.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(*x, *y, box_w, box_h)
                    .font_size(18.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }

        // Connector 1: Straight — Plan to Design (right side of box 1 to left side of box 2)
        let (ct, co) = ConnectorOptionsBuilder::new()
            .connector_type(ConnectorType::Straight)
            .x1(1.0 + box_w)
            .y1(1.3 + box_h / 2.0)
            .x2(5.0)
            .y2(1.3 + box_h / 2.0)
            .line(ShapeLineProps {
                color: Some("4472C4".into()),
                width: Some(2.5),
                dash_type: None,
                transparency: None,
                begin_arrow_type: None,
                end_arrow_type: Some("triangle".into()),
                cap: None,
                join: None,
            })
            .build();
        slide.add_connector(ct, co);

        // Label
        slide.add_text(
            "Straight",
            TextOptionsBuilder::new()
                .bounds(3.5, 0.9, 2.0, 0.4)
                .font_size(10.0)
                .color("#4472C4")
                .align(AlignH::Center)
                .build(),
        );

        // Connector 2: Elbow — Design down to Build (bottom of box 2 across to top of box 3)
        let (ct, co) = ConnectorOptionsBuilder::new()
            .connector_type(ConnectorType::Elbow)
            .x1(5.0 + box_w / 2.0)
            .y1(1.3 + box_h)
            .x2(1.0 + box_w / 2.0)
            .y2(3.5)
            .line(ShapeLineProps {
                color: Some("ED7D31".into()),
                width: Some(3.0),
                dash_type: Some("dash".into()),
                transparency: None,
                begin_arrow_type: None,
                end_arrow_type: Some("triangle".into()),
                cap: None,
                join: None,
            })
            .build();
        slide.add_connector(ct, co);

        // Label
        slide.add_text(
            "Elbow (dashed)",
            TextOptionsBuilder::new()
                .bounds(3.2, 2.55, 2.6, 0.4)
                .font_size(10.0)
                .color("#ED7D31")
                .align(AlignH::Center)
                .build(),
        );

        // Connector 3: Curved — Build to Deploy
        let (ct, co) = ConnectorOptionsBuilder::new()
            .connector_type(ConnectorType::Curved)
            .x1(1.0 + box_w)
            .y1(3.5 + box_h / 2.0)
            .x2(5.0)
            .y2(3.5 + box_h / 2.0)
            .line(ShapeLineProps {
                color: Some("70AD47".into()),
                width: Some(2.0),
                dash_type: Some("lgDash".into()),
                transparency: None,
                begin_arrow_type: None,
                end_arrow_type: Some("triangle".into()),
                cap: None,
                join: None,
            })
            .build();
        slide.add_connector(ct, co);

        // Label
        slide.add_text(
            "Curved (long dash)",
            TextOptionsBuilder::new()
                .bounds(3.0, 4.8, 3.0, 0.4)
                .font_size(10.0)
                .color("#70AD47")
                .align(AlignH::Center)
                .build(),
        );
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: Hub-and-spoke network diagram with elbow connectors
    // ══════════════════════════════════════════════════════════
    {
        let slide = pres.add_slide();
        slide.set_background_color("#1B2838");

        slide.add_text(
            "Hub-and-Spoke Network",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.5)
                .font_size(22.0)
                .bold()
                .color("#FFFFFF")
                .build(),
        );

        // Central hub
        let hub_x = 3.75;
        let hub_y = 2.1;
        let hub_w = 2.5;
        let hub_h = 1.5;

        slide.add_shape(
            ShapeType::Ellipse,
            ShapeOptionsBuilder::new()
                .bounds(hub_x, hub_y, hub_w, hub_h)
                .fill_color("#E74C3C")
                .build(),
        );
        slide.add_text(
            "Core\nServer",
            TextOptionsBuilder::new()
                .bounds(hub_x, hub_y, hub_w, hub_h)
                .font_size(16.0)
                .bold()
                .color("#FFFFFF")
                .align(AlignH::Center)
                .valign(AlignV::Middle)
                .build(),
        );

        // Spoke nodes arranged around hub
        let spokes: [(&str, f64, f64, &str); 5] = [
            ("Web App",    1.0,  0.7, "#3498DB"),
            ("Mobile",     7.0,  0.7, "#2ECC71"),
            ("Database",   0.3,  3.8, "#9B59B6"),
            ("Cache",      7.5,  3.8, "#F39C12"),
            ("CDN",        4.0,  4.6, "#1ABC9C"),
        ];
        let spoke_w = 1.8;
        let spoke_h = 1.0;

        let hub_cx = hub_x + hub_w / 2.0;
        let hub_cy = hub_y + hub_h / 2.0;

        for (label, sx, sy, color) in &spokes {
            slide.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .bounds(*sx, *sy, spoke_w, spoke_h)
                    .fill_color(*color)
                    .rect_radius(0.1)
                    .build(),
            );
            slide.add_text(
                *label,
                TextOptionsBuilder::new()
                    .bounds(*sx, *sy, spoke_w, spoke_h)
                    .font_size(13.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );

            // Elbow connector from hub center to spoke center
            let spoke_cx = sx + spoke_w / 2.0;
            let spoke_cy = sy + spoke_h / 2.0;

            let (ct, co) = ConnectorOptionsBuilder::new()
                .connector_type(ConnectorType::Elbow)
                .x1(hub_cx)
                .y1(hub_cy)
                .x2(spoke_cx)
                .y2(spoke_cy)
                .line(ShapeLineProps {
                    color: Some(color.trim_start_matches('#').into()),
                    width: Some(1.5),
                    dash_type: None,
                    transparency: Some(20.0),
                    begin_arrow_type: None,
                    end_arrow_type: Some("triangle".into()),
                    cap: None,
                    join: None,
                })
                .build();
            slide.add_connector(ct, co);
        }
    }

    pres.write_to_file("15_connectors.pptx").unwrap();
    println!("Wrote 15_connectors.pptx");
}
