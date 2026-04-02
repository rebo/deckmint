//! Shadows and 3D effects showcase: outer/inner shadows with various configs,
//! bevel presets, camera presets, and light rigs on shapes.

use deckmint::layout::GridLayoutBuilder;
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::types::*;
use deckmint::{AlignH, AlignV, Presentation, ShapeType};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Shadows & 3D Effects".to_string();

    // ══════════════════════════════════════════════════════════
    // Slide 1: Shadow types and configurations
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#F0F2F5");

        s.add_text(
            "Shadow Effects Gallery",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.6)
                .font_size(28.0)
                .bold()
                .color("#1B2A4A")
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(3, 2, 0.3)
            .origin(0.5, 1.0)
            .container(9.0, 4.3)
            .build();

        let shadow_configs: Vec<(&str, &str, ShadowProps, &str)> = vec![
            (
                "Soft Drop",
                "#4472C4",
                ShadowProps::outer().with_blur(8.0).with_offset(4.0).with_opacity(0.4),
                "blur=8, offset=4",
            ),
            (
                "Sharp Drop",
                "#ED7D31",
                ShadowProps::outer().with_blur(1.0).with_offset(3.0).with_opacity(0.6),
                "blur=1, offset=3",
            ),
            (
                "Long Shadow",
                "#70AD47",
                ShadowProps::outer()
                    .with_blur(4.0)
                    .with_offset(8.0)
                    .with_angle(135.0)
                    .with_opacity(0.35),
                "offset=8, angle=135",
            ),
            (
                "Inner Shadow",
                "#5B9BD5",
                ShadowProps::inner().with_blur(6.0).with_color("#000000").with_opacity(0.5),
                "inner, blur=6",
            ),
            (
                "Colored Shadow",
                "#FFC000",
                ShadowProps::outer()
                    .with_blur(10.0)
                    .with_offset(3.0)
                    .with_color("#FF6B00")
                    .with_opacity(0.6),
                "color=#FF6B00",
            ),
            (
                "Inner Colored",
                "#9B59B6",
                ShadowProps::inner()
                    .with_blur(8.0)
                    .with_color("#4A0080")
                    .with_opacity(0.7),
                "inner, color=#4A0080",
            ),
        ];

        for (i, (label, fill, shadow, desc)) in shadow_configs.iter().enumerate() {
            let col = i % 3;
            let row = i / 3;
            let cell = grid.cell(col, row);

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(cell.inset(0.15))
                    .fill_color(*fill)
                    .shadow(shadow.clone())
                    .rect_radius(0.12)
                    .build(),
            );

            let inner = cell.inset(0.15);
            let (top, bot) = inner.halves_v(0.0);
            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .rect(top)
                    .font_size(18.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Bottom)
                    .build(),
            );
            s.add_text(
                *desc,
                TextOptionsBuilder::new()
                    .rect(bot)
                    .font_size(11.0)
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Top)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 2: 3D bevel presets
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#E8ECF1");

        s.add_text(
            "3D Bevel Presets",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.15, 9.0, 0.55)
                .font_size(28.0)
                .bold()
                .color("#1B2A4A")
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(4, 3, 0.2)
            .origin(0.3, 0.8)
            .container(9.4, 4.6)
            .build();

        let bevels: Vec<(&str, BevelPreset, &str)> = vec![
            ("Circle", BevelPreset::Circle, "#4472C4"),
            ("Relaxed Inset", BevelPreset::RelaxedInset, "#ED7D31"),
            ("Angle", BevelPreset::Angle, "#70AD47"),
            ("Soft Round", BevelPreset::SoftRound, "#FFC000"),
            ("Convex", BevelPreset::Convex, "#5B9BD5"),
            ("Slope", BevelPreset::Slope, "#9B59B6"),
            ("Divot", BevelPreset::Divot, "#E74C3C"),
            ("Riblet", BevelPreset::Riblet, "#2ECC71"),
            ("Hard Edge", BevelPreset::HardEdge, "#3498DB"),
            ("Art Deco", BevelPreset::ArtDeco, "#F39C12"),
            ("Cross", BevelPreset::Cross, "#1ABC9C"),
            ("Cool Slant", BevelPreset::CoolSlant, "#E67E22"),
        ];

        let standard_scene = Scene3DProps {
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
        };

        for (i, (label, preset, color)) in bevels.iter().enumerate() {
            let col = i % 4;
            let row = i / 4;
            let cell = grid.cell(col, row);
            let shape_rect = cell.inset(0.08);

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(shape_rect)
                    .fill_color(*color)
                    .rect_radius(0.08)
                    .shape_3d(Shape3DProps {
                        bevel_top: Some(BevelProps::new(*preset).with_size(63500, 25400)),
                        bevel_bottom: None,
                        extrusion_height: Some(25400),
                        contour_width: None,
                        contour_color: None,
                        material: Some(MaterialPreset::WarmMatte),
                    })
                    .scene_3d(standard_scene.clone())
                    .build(),
            );

            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .rect(shape_rect)
                    .font_size(12.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }
    }

    // ══════════════════════════════════════════════════════════
    // Slide 3: Camera presets and light rigs
    // ══════════════════════════════════════════════════════════
    {
        let s = pres.add_slide();
        s.set_background_color("#1B2A4A");

        s.add_text(
            "3D Camera & Lighting",
            TextOptionsBuilder::new()
                .bounds(0.5, 0.1, 9.0, 0.55)
                .font_size(28.0)
                .bold()
                .color("#FFFFFF")
                .build(),
        );

        let grid = GridLayoutBuilder::grid_n_m(3, 3, 0.2)
            .origin(0.3, 0.75)
            .container(9.4, 4.7)
            .build();

        let camera_configs: Vec<(&str, CameraPreset, LightRigType, LightDirection, &str)> = vec![
            ("Ortho Front\nThree Point", CameraPreset::OrthographicFront, LightRigType::ThreePt, LightDirection::Top, "#4472C4"),
            ("Perspective\nBalanced", CameraPreset::PerspectiveFront, LightRigType::Balanced, LightDirection::TopRight, "#ED7D31"),
            ("Iso Top Up\nHarsh", CameraPreset::IsometricTopUp, LightRigType::Harsh, LightDirection::Top, "#70AD47"),
            ("Iso Top Down\nFlood", CameraPreset::IsometricTopDown, LightRigType::Flood, LightDirection::Left, "#FFC000"),
            ("Iso Left Down\nSoft", CameraPreset::IsometricLeftDown, LightRigType::Soft, LightDirection::TopLeft, "#5B9BD5"),
            ("Iso Right Up\nMorning", CameraPreset::IsometricRightUp, LightRigType::Morning, LightDirection::Right, "#9B59B6"),
            ("Oblique Top-L\nSunrise", CameraPreset::ObliqueTopLeft, LightRigType::Sunrise, LightDirection::TopRight, "#E74C3C"),
            ("Oblique Top\nContrasting", CameraPreset::ObliqueTop, LightRigType::Contrasting, LightDirection::Top, "#2ECC71"),
            ("Oblique Top-R\nGlow", CameraPreset::ObliqueTopRight, LightRigType::Glow, LightDirection::Bottom, "#F39C12"),
        ];

        for (i, (label, cam_preset, light_type, light_dir, color)) in camera_configs.iter().enumerate() {
            let col = i % 3;
            let row = i / 3;
            let cell = grid.cell(col, row);
            let shape_rect = cell.inset(0.1);

            s.add_shape(
                ShapeType::RoundRect,
                ShapeOptionsBuilder::new()
                    .rect(shape_rect)
                    .fill_color(*color)
                    .rect_radius(0.1)
                    .shape_3d(Shape3DProps {
                        bevel_top: Some(BevelProps::new(BevelPreset::Circle).with_size(76200, 38100)),
                        bevel_bottom: Some(BevelProps::new(BevelPreset::Circle)),
                        extrusion_height: Some(50800),
                        contour_width: Some(12700),
                        contour_color: Some("FFFFFF".to_string()),
                        material: Some(MaterialPreset::Metal),
                    })
                    .scene_3d(Scene3DProps {
                        camera: Camera3D {
                            preset: *cam_preset,
                            fov: None,
                            rotation: Some(Rotation3D::from_degrees(0.0, 0.0, 0.0)),
                        },
                        light_rig: LightRig3D {
                            rig_type: *light_type,
                            direction: *light_dir,
                            rotation: None,
                        },
                    })
                    .build(),
            );

            s.add_text(
                *label,
                TextOptionsBuilder::new()
                    .rect(shape_rect)
                    .font_size(11.0)
                    .bold()
                    .color("#FFFFFF")
                    .align(AlignH::Center)
                    .valign(AlignV::Middle)
                    .build(),
            );
        }
    }

    pres.write_to_file("26_shadows_3d.pptx").unwrap();
    println!("Wrote 26_shadows_3d.pptx");
}
