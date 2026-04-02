# deckmint

Create PowerPoint `.pptx` presentations programmatically in Rust.

## Quick start

```rust
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
std::fs::write("hello.pptx", bytes).unwrap();
```

## Features

- **Text** -- rich text runs, bold/italic/underline (15 styles), alignment, RTL, bullets, superscript/subscript, glow, outline, hyperlinks, tab stops, text fit/wrap
- **Shapes** -- 150+ preset shapes, fill, line, shadow, rotation, flip, rounded corners, arrow heads, custom geometry
- **Images** -- PNG/JPEG/GIF/SVG, cover/contain/crop sizing, transparency, rounding, auto-dimension detection, base64 input
- **Tables** -- column widths, per-row heights, cell fill/color/font, alignment, margins, borders, colspan/rowspan, hyperlinks
- **Charts** -- bar, column, line, area, pie, doughnut, scatter with data labels, legends, axis configuration, gridlines
- **Animations** -- entrance/exit/emphasis effects with triggers, delays, and sequencing
- **Slide masters** -- background color/image, reusable objects, custom layouts
- **Sections** -- organize slides into named sections
- **Transitions** -- fade, push, wipe, and more with speed control
- **Connectors** -- straight, elbow, and curved connectors between shapes
- **Groups** -- group multiple objects together
- **Media** -- embed video and audio
- **Equations** -- LaTeX math via the `deckmint-math` crate (enabled by default)
- **Grid layout** -- arrange objects in rows/columns with gutters and padding
- **Speaker notes** -- add notes to any slide
- **WASM** -- full browser support via the `deckmint-wasm` crate

## Crates

| Crate | Description |
|---|---|
| [`deckmint`](deckmint/) | Core library |
| [`deckmint-math`](deckmint-math/) | LaTeX/MathML to OMML equation converter |
| [`deckmint-wasm`](deckmint-wasm/) | WebAssembly bindings |

## Examples

```bash
cargo run --example 01_hello_world
cargo run --example features
```

See the [`deckmint/examples/`](deckmint/examples/) directory for all examples.

## License

MIT

## Acknowledgements

This project has learned from and been inspired by:

- [PptxGenJS](https://github.com/gitbrent/PptxGenJS) -- the original JavaScript PowerPoint generation library
- [python-pptx](https://github.com/scanny/python-pptx) -- Python library for creating PowerPoint files
- [rust-pptx](https://github.com/cstkingkey/rust-pptx) -- Rust PPTX reading/writing crate
