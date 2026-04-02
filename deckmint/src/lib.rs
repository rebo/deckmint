/*!
# deckmint

Create PowerPoint `.pptx` presentations programmatically in Rust.
Works on native targets and compiles to WASM via the companion `deckmint-wasm` crate.

## Quick start

```rust,no_run
use deckmint::{Presentation, ShapeType};
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::objects::shape::ShapeOptionsBuilder;

let mut pres = Presentation::new();
pres.title = "My Presentation".to_string();

let slide = pres.add_slide();
slide.add_text(
    "Hello, World!",
    TextOptionsBuilder::new()
        .x(1.0).y(1.5).w(8.0).h(1.5)
        .font_size(36.0)
        .bold()
        .build()
);
slide.add_shape(
    ShapeType::Rect,
    ShapeOptionsBuilder::new()
        .x(1.0).y(3.0).w(4.0).h(2.0)
        .fill_color("4472C4")
        .build()
);

let bytes = pres.write().expect("failed to generate pptx");
// bytes is a valid .pptx ZIP archive
```

## Tables

```rust,no_run
use deckmint::Presentation;
use deckmint::objects::table::{TableOptionsBuilder, TableCell};

let mut pres = Presentation::new();
let slide = pres.add_slide();
slide.add_table(
    vec![
        vec![TableCell::new("Name"), TableCell::new("Score")],
        vec![TableCell::new("Alice"), TableCell::new("95")],
    ],
    TableOptionsBuilder::new()
        .x(1.0).y(1.5).w(8.0)
        .col_w(vec![4.0, 4.0])
        .font_size(14.0)
        .build(),
);
```

## Charts

```rust,no_run
use deckmint::{Presentation, ChartType, ChartOptionsBuilder, ChartSeries};

let mut pres = Presentation::new();
let slide = pres.add_slide();
let series = vec![
    ChartSeries::new("Q1", vec!["Jan", "Feb", "Mar"], vec![10.0, 20.0, 30.0]),
];
slide.add_chart(
    ChartType::Bar,
    series,
    ChartOptionsBuilder::new()
        .x(0.5).y(1.0).w(9.0).h(4.5)
        .title("Monthly Sales")
        .show_value()
        .build(),
);
```

## Images

```rust,no_run
use deckmint::Presentation;
use deckmint::objects::image::ImageOptionsBuilder;

let png_bytes: Vec<u8> = std::fs::read("logo.png").unwrap();
let mut pres = Presentation::new();
let slide = pres.add_slide();
slide.add_image_from(
    ImageOptionsBuilder::new()
        .bytes(png_bytes, "png")
        .pos(3.0, 2.0).size(4.0, 3.0)
).unwrap();
```

## Animations

```rust,no_run
use deckmint::{Presentation, AnimationEffect, Direction};
use deckmint::objects::text::TextOptionsBuilder;

let mut pres = Presentation::new();
let slide = pres.add_slide();
slide.add_text(
    "I fly in!",
    TextOptionsBuilder::new()
        .x(1.0).y(2.0).w(8.0).h(1.5)
        .animation(AnimationEffect::fly_in(Direction::Left))
        .build(),
);
```

## Equations (requires `math` feature, enabled by default)

```rust,no_run
use deckmint::Presentation;
use deckmint::objects::text::TextOptionsBuilder;

let mut pres = Presentation::new();
let slide = pres.add_slide();

// Standalone equation
slide.add_equation(
    r"x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}",
    TextOptionsBuilder::new().pos(1.0, 1.5).size(8.0, 1.0).build(),
).unwrap();

// Inline math mixed with text (delimited by $...$)
slide.add_text_with_math(
    r"The area of a circle is $A = \pi r^2$.",
    TextOptionsBuilder::new().pos(1.0, 3.0).size(8.0, 1.0).font_size(18.0).build(),
).unwrap();
```

## Grid layout

The [`layout`] module provides a [`layout::GridLayoutBuilder`] for
arranging objects in rows and columns with gutters and padding.

## Output

- [`Presentation::write()`] returns `Vec<u8>` (the `.pptx` ZIP archive).
- [`Presentation::write_to_file(path)`](Presentation::write_to_file) writes directly to disk.
*/

pub mod enums;
pub mod error;
pub mod layout;
pub mod objects;
pub mod packaging;
pub mod presentation;
pub mod slide;
pub mod types;
pub mod utils;
pub mod validate;
pub mod xml;

// Re-exports for ergonomic top-level access
pub use enums::{AlignH, AlignV, ChartType, SchemeColor, ShapeType};
pub use error::PptxError;
pub use presentation::{Presentation, SectionDef, SlideLayout, SlideMasterDef, ThemeProps};
pub use slide::Slide;
pub use types::{
    AnimationEffect, AnimationEffectType, AnimationTrigger, TextTarget,
    CheckerboardDir, CustomGeomPoint, Direction, ShapeVariant, SplitOrientation, StripDir,
    BorderProps, BorderType, Color, Coord, FieldType, FillType, GlowProps, GradientFill,
    GradientStop, GradientType, HyperlinkAction, HyperlinkProps, ImageColorAdjust, Margin,
    PatternFill, PatternType, PositionProps, PresLayout, ShadowProps, ShadowType,
    BevelPreset, BevelProps, Camera3D, CameraPreset, LightDirection, LightRig3D, LightRigType,
    LineCap, LineJoin, MaterialPreset, Rotation3D, Scene3DProps, Shape3DProps,
    ShapeFillProps, ShapeLineProps, SlideNumberProps, TextOutlineProps, ThemeColorMod,
    TransitionDir, TransitionProps, TransitionSpeed, TransitionType,
};
pub use objects::chart::{BarDir, BarGrouping, ChartOptions, ChartOptionsBuilder, ChartSeries, LegendPos};
pub use objects::text::{char_range_for_run, TabStop, TextFit, TextRunBuilder};
pub use objects::connector::{ConnectorEndpoint, ConnectorOptions, ConnectorOptionsBuilder, ConnectorType};
pub use objects::media::{MediaOptions, MediaOptionsBuilder, MediaType};
pub use objects::group::{GroupOptions, GroupOptionsBuilder};
pub use layout::{
    CellRect, GridLayout, GridLayoutBuilder, GridTrack,
    center_in, align_left, align_right, align_top, align_bottom, split_h, split_v,
};
