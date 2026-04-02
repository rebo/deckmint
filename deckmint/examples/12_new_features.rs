/// Example demonstrating all new features:
/// - Line cap & join styles
/// - Image cropping
/// - Hover actions
/// - Group shapes
/// - 3D effects
/// - Bubble, stock, and surface charts
/// - Validation layer

fn main() {
    use deckmint::*;
    use deckmint::objects::text::TextOptionsBuilder;
    use deckmint::objects::shape::ShapeOptionsBuilder;
    use deckmint::objects::image::ImageOptionsBuilder;

    // Load tiger image from disk
    let tiger_png = std::fs::read(concat!(env!("CARGO_MANIFEST_DIR"), "/examples/tiger.png"))
        .expect("tiger.png not found — place it in deckmint/examples/");

    let mut pres = Presentation::new();

    // ── Slide 1: Line cap & join styles ─────────────
    {
        let slide = pres.add_slide();
        slide.add_text("Line Cap & Join Styles", TextOptionsBuilder::new()
            .pos(0.5, 0.3).size(9.0, 0.6).font_size(24.0).bold().build());

        // Flat cap, miter join
        slide.add_shape(ShapeType::Line, ShapeOptionsBuilder::new()
            .pos(1.0, 1.5).size(3.0, 0.0)
            .line_color("FF0000").line_width(6.0)
            .line_cap(LineCap::Flat).line_join(LineJoin::Miter)
            .build());
        slide.add_text("Flat cap, Miter join", TextOptionsBuilder::new()
            .pos(1.0, 1.1).size(3.0, 0.4).font_size(11.0).build());

        // Round cap, round join
        slide.add_shape(ShapeType::Line, ShapeOptionsBuilder::new()
            .pos(1.0, 2.8).size(3.0, 0.0)
            .line_color("00AA00").line_width(6.0)
            .line_cap(LineCap::Round).line_join(LineJoin::Round)
            .build());
        slide.add_text("Round cap, Round join", TextOptionsBuilder::new()
            .pos(1.0, 2.4).size(3.0, 0.4).font_size(11.0).build());

        // Square cap, bevel join
        slide.add_shape(ShapeType::Line, ShapeOptionsBuilder::new()
            .pos(1.0, 4.1).size(3.0, 0.0)
            .line_color("0000FF").line_width(6.0)
            .line_cap(LineCap::Square).line_join(LineJoin::Bevel)
            .build());
        slide.add_text("Square cap, Bevel join", TextOptionsBuilder::new()
            .pos(1.0, 3.7).size(3.0, 0.4).font_size(11.0).build());
    }

    // ── Slide 2: Image cropping ─────────────────────
    {
        let slide = pres.add_slide();
        slide.add_text("Image Cropping (LTRB)", TextOptionsBuilder::new()
            .pos(0.5, 0.3).size(9.0, 0.6).font_size(24.0).bold().build());

        // Original (no crop)
        slide.add_text("Original", TextOptionsBuilder::new()
            .pos(0.5, 1.2).size(2.0, 0.3).font_size(11.0).build());
        slide.add_image(tiger_png.clone(), "png", ImageOptionsBuilder::new()
            .pos(0.5, 1.6).size(2.5, 3.3).build_opts());

        // Cropped 20% from each side
        slide.add_text("Crop 20% each side", TextOptionsBuilder::new()
            .pos(4.0, 1.2).size(2.5, 0.3).font_size(11.0).build());
        slide.add_image(tiger_png.clone(), "png", ImageOptionsBuilder::new()
            .pos(4.0, 1.6).size(2.5, 3.3)
            .crop(0.2, 0.2, 0.2, 0.2).build_opts());
    }

    // ── Slide 3: 3D effects ─────────────────────────
    {
        let slide = pres.add_slide();
        slide.add_text("3D Effects on Shapes", TextOptionsBuilder::new()
            .pos(0.5, 0.3).size(9.0, 0.6).font_size(24.0).bold().build());

        // Shape with bevel and material
        slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
            .pos(1.0, 1.5).size(3.0, 2.0)
            .fill_color("4472C4")
            .shape_3d(Shape3DProps {
                bevel_top: Some(BevelProps::new(BevelPreset::Circle).with_size(63500, 25400)),
                bevel_bottom: None,
                extrusion_height: Some(76200),
                contour_width: Some(12700),
                contour_color: Some("2F5597".to_string()),
                material: Some(MaterialPreset::WarmMatte),
            })
            .scene_3d(Scene3DProps {
                camera: Camera3D {
                    preset: CameraPreset::OrthographicFront,
                    fov: None,
                    rotation: None,
                },
                light_rig: LightRig3D {
                    rig_type: LightRigType::ThreePt,
                    direction: LightDirection::Top,
                    rotation: None,
                },
            })
            .build());
        slide.add_text("Circle bevel + WarmMatte", TextOptionsBuilder::new()
            .pos(1.0, 3.7).size(3.0, 0.4).font_size(11.0).build());

        // Shape with different bevel
        slide.add_shape(ShapeType::Hexagon, ShapeOptionsBuilder::new()
            .pos(5.5, 1.5).size(3.0, 2.0)
            .fill_color("ED7D31")
            .shape_3d(Shape3DProps {
                bevel_top: Some(BevelProps::new(BevelPreset::ArtDeco).with_size(50800, 50800)),
                bevel_bottom: Some(BevelProps::new(BevelPreset::Angle)),
                extrusion_height: None,
                contour_width: None,
                contour_color: None,
                material: Some(MaterialPreset::Metal),
            })
            .scene_3d(Scene3DProps {
                camera: Camera3D {
                    preset: CameraPreset::IsometricTopUp,
                    fov: None,
                    rotation: Some(Rotation3D::from_degrees(0.0, 0.0, 0.0)),
                },
                light_rig: LightRig3D {
                    rig_type: LightRigType::Balanced,
                    direction: LightDirection::TopRight,
                    rotation: None,
                },
            })
            .build());
        slide.add_text("ArtDeco bevel + Metal", TextOptionsBuilder::new()
            .pos(5.5, 3.7).size(3.0, 0.4).font_size(11.0).build());
    }

    // ── Slide 4: Group shapes ───────────────────────
    {
        let slide = pres.add_slide();
        slide.add_text("Group Shapes", TextOptionsBuilder::new()
            .pos(0.5, 0.3).size(9.0, 0.6).font_size(24.0).bold().build());

        // Create child shapes
        use deckmint::objects::SlideObject;
        let child1 = SlideObject::Shape(deckmint::objects::ShapeObject {
            object_name: "Child1".to_string(),
            shape_type: ShapeType::Rect,
            options: ShapeOptionsBuilder::new()
                .pos(0.0, 0.0).size(2.0, 1.5)
                .fill_color("4472C4").build(),
            text: None,
        });
        let child2 = SlideObject::Shape(deckmint::objects::ShapeObject {
            object_name: "Child2".to_string(),
            shape_type: ShapeType::Ellipse,
            options: ShapeOptionsBuilder::new()
                .pos(2.5, 0.5).size(1.5, 1.0)
                .fill_color("ED7D31").build(),
            text: None,
        });

        slide.add_group(vec![child1, child2], GroupOptionsBuilder::new()
            .pos(2.0, 1.5).size(5.0, 3.0).build());
    }

    // ── Slide 5: Bubble chart ───────────────────────
    {
        let slide = pres.add_slide();
        slide.add_text("Bubble Chart", TextOptionsBuilder::new()
            .pos(0.5, 0.3).size(9.0, 0.6).font_size(24.0).bold().build());

        let series = vec![
            ChartSeries::new("Revenue", vec!["Q1", "Q2", "Q3", "Q4"], vec![10.0, 25.0, 15.0, 30.0])
                .sizes(vec![5.0, 10.0, 7.0, 15.0])
                .color("4472C4"),
            ChartSeries::new("Profit", vec!["Q1", "Q2", "Q3", "Q4"], vec![5.0, 12.0, 8.0, 20.0])
                .sizes(vec![3.0, 8.0, 5.0, 12.0])
                .color("ED7D31"),
        ];
        slide.add_chart(ChartType::Bubble, series, ChartOptionsBuilder::new()
            .pos(0.5, 1.0).size(9.0, 5.5)
            .title("Revenue vs Profit").show_legend(true).build());
    }

    // ── Slide 6: Stock chart (OHLC) ─────────────────
    {
        let slide = pres.add_slide();
        slide.add_text("Stock Chart (OHLC)", TextOptionsBuilder::new()
            .pos(0.5, 0.3).size(9.0, 0.6).font_size(24.0).bold().build());

        let days: Vec<String> = vec!["Mon", "Tue", "Wed", "Thu", "Fri"].into_iter().map(String::from).collect();
        let series = vec![
            ChartSeries::new("Open",  days.clone(), vec![100.0, 105.0, 102.0, 108.0, 106.0]),
            ChartSeries::new("High",  days.clone(), vec![110.0, 112.0, 109.0, 115.0, 113.0]),
            ChartSeries::new("Low",   days.clone(), vec![95.0, 100.0, 98.0, 103.0, 101.0]),
            ChartSeries::new("Close", days,         vec![105.0, 102.0, 108.0, 106.0, 110.0]),
        ];
        slide.add_chart(ChartType::StockOHLC, series, ChartOptionsBuilder::new()
            .pos(0.5, 1.0).size(9.0, 5.5)
            .title("Weekly Stock Prices").build());
    }

    // ── Slide 7: Surface chart ──────────────────────
    {
        let slide = pres.add_slide();
        slide.add_text("Surface Chart", TextOptionsBuilder::new()
            .pos(0.5, 0.3).size(9.0, 0.6).font_size(24.0).bold().build());

        let cats: Vec<String> = vec!["A", "B", "C", "D"].into_iter().map(String::from).collect();
        let series = vec![
            ChartSeries::new("Series 1", cats.clone(), vec![1.0, 3.0, 2.0, 4.0]).color("4472C4"),
            ChartSeries::new("Series 2", cats.clone(), vec![2.0, 4.0, 1.0, 3.0]).color("ED7D31"),
            ChartSeries::new("Series 3", cats,         vec![3.0, 1.0, 4.0, 2.0]).color("A5A5A5"),
        ];
        slide.add_chart(ChartType::Surface, series, ChartOptionsBuilder::new()
            .pos(0.5, 1.0).size(9.0, 5.5)
            .title("3D Surface").show_legend(true).build());
    }

    // ── Write with validation ───────────────────────
    match pres.write_validated() {
        Ok(bytes) => {
            std::fs::write("12_new_features.pptx", &bytes).expect("write file");
            println!("Wrote 12_new_features.pptx (validated, {} bytes)", bytes.len());
        }
        Err(e) => {
            eprintln!("Validation failed: {e}");
            // Fall back to writing without validation
            pres.write_to_file("12_new_features.pptx").expect("write file");
            println!("Wrote 12_new_features.pptx (unvalidated)");
        }
    }
}

