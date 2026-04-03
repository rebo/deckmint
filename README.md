# deckmint

Create PowerPoint `.pptx` presentations programmatically in Rust. Works on native targets and compiles to WebAssembly for in-browser use.

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
        .bounds(1.0, 1.5, 8.0, 1.5)
        .font_size(36.0)
        .bold()
        .color("#4472C4")
        .build()
);
slide.add_shape(
    ShapeType::Rect,
    ShapeOptionsBuilder::new()
        .bounds(1.0, 3.5, 4.0, 1.5)
        .fill_color("#4472C4")
        .build()
);

std::fs::write("hello.pptx", pres.write().unwrap()).unwrap();
```

## Equations

LaTeX equations are converted to native, editable Office Math (OMML). Enabled by default via the `deckmint-math` crate.

```rust
use deckmint::Presentation;
use deckmint::objects::text::TextOptionsBuilder;

let mut pres = Presentation::new();
let slide = pres.add_slide();

slide.add_equation(
    r"\int_{0}^{\infty} e^{-x^2} \, dx = \frac{\sqrt{\pi}}{2}",
    TextOptionsBuilder::new().bounds(1.0, 2.0, 8.0, 1.5).font_size(28.0).build(),
).unwrap();
```

## Grid layout

The built-in grid layout system provides CSS-grid-inspired positioning with flexible/fixed tracks, gaps, and nesting -- no manual coordinate math needed.

```rust
use deckmint::{Presentation, ShapeType, AlignH, AlignV};
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::layout::{GridLayoutBuilder, GridTrack};

let mut pres = Presentation::new();
let slide = pres.add_slide();

// Header + flexible content area + footer
let grid = GridLayoutBuilder::new()
    .cols(vec![GridTrack::Fr(1.0)])
    .rows(vec![
        GridTrack::Inches(0.8),   // header
        GridTrack::Fr(1.0),       // content
        GridTrack::Inches(0.4),   // footer
    ])
    .row_gap(0.1)
    .padding(0.3)
    .build();

slide.add_shape(ShapeType::Rect, ShapeOptionsBuilder::new()
    .rect(grid.cell(0, 0)).fill_color("#1B2A4A").build());
slide.add_text("Dashboard", TextOptionsBuilder::new()
    .rect(grid.cell(0, 0))
    .font_size(24.0).bold().color("#FFFFFF")
    .align(AlignH::Center).valign(AlignV::Middle)
    .build());

// Nest a 3-column grid inside the content cell
let cards = grid.sub_grid(0, 1)
    .cols(vec![GridTrack::Fr(1.0); 3])
    .gap(0.15)
    .build();

for (col, (label, color)) in [("Users", "#4472C4"), ("Revenue", "#70AD47"), ("Growth", "#ED7D31")]
    .iter().enumerate()
{
    let cell = cards.cell(col, 0);
    slide.add_shape(ShapeType::RoundRect, ShapeOptionsBuilder::new()
        .rect(cell).fill_color(color).rect_radius(0.08).build());
    slide.add_text(label, TextOptionsBuilder::new()
        .rect(cell).font_size(18.0).bold().color("#FFFFFF")
        .align(AlignH::Center).valign(AlignV::Middle).build());
}
```

## WebAssembly

deckmint compiles to WASM via the `deckmint-wasm` crate, enabling PowerPoint generation directly in the browser with no server round-trip. See [`deckmint-wasm/demo/`](deckmint-wasm/demo/) for a working example that generates and downloads a `.pptx` from a button click.

```bash
# Build the WASM package and run the demo
cd deckmint-wasm
wasm-pack build --target web
python3 -m http.server 8080
# Open http://localhost:8080/demo/
```

## Features

- **Text** -- rich text runs, bold/italic/underline (15 styles), alignment, RTL, bullets, superscript/subscript, glow, outline, hyperlinks, tab stops, text fit/wrap
- **Shapes** -- 150+ preset shapes, fill, line, shadow, rotation, flip, rounded corners, arrow heads, custom geometry
- **Images** -- PNG/JPEG/GIF/SVG, cover/contain/crop sizing, transparency, rounding, auto-dimension detection, base64 input
- **Tables** -- column widths, per-row heights, cell fill/color/font, alignment, margins, borders, colspan/rowspan, hyperlinks
- **Charts** -- bar, column, line, area, pie, doughnut, scatter, radar, bubble, stock, surface with data labels, legends, axis configuration
- **Animations** -- entrance/exit/emphasis effects with triggers, delays, sequencing, and click groups
- **Slide masters** -- background color/image, reusable objects, custom layouts, promote-from-slide
- **Transitions** -- 30+ types (fade, push, wipe, zoom, flash, morph, etc.) with speed and direction control
- **Equations** -- LaTeX math converted to native editable OMML via `deckmint-math`
- **Grid layout** -- CSS-grid-inspired positioning with `Fr`/`Inches`/`Percent`/`MinMax` tracks, gaps, padding, nesting
- **Connectors** -- straight, elbow, and curved connectors between shapes
- **Groups** -- group multiple objects together
- **Media** -- embed video and audio
- **Sections** -- organize slides into named sections
- **Speaker notes** -- add notes to any slide
- **3D effects** -- bevel presets, materials, camera angles, light rigs
- **Pattern fills** -- 43 pattern types with foreground/background colors
- **WASM** -- full browser support via `deckmint-wasm`

## Crates

| Crate | Description |
|---|---|
| [`deckmint`](deckmint/) | Core library |
| [`deckmint-math`](deckmint-math/) | LaTeX/MathML to OMML equation converter |
| [`deckmint-wasm`](deckmint-wasm/) | WebAssembly bindings |

## Examples

There are 30+ examples covering every feature:

```bash
cargo run --example 01_hello_world    # basics
cargo run --example 09_grid_layout    # grid positioning
cargo run --example 11_equations      # LaTeX math
cargo run --example 13_master_slides  # slide masters & promotion
cargo run --example 30_pitch_deck     # full pitch deck
cargo run --example features          # comprehensive feature demo
```

See [`deckmint/examples/`](deckmint/examples/) for the full list.

## License

MIT

## Acknowledgements

This project has learned from and been inspired by:

- [PptxGenJS](https://github.com/gitbrent/PptxGenJS) -- the original JavaScript PowerPoint generation library
- [python-pptx](https://github.com/scanny/python-pptx) -- Python library for creating PowerPoint files
- [rust-pptx](https://github.com/cstkingkey/rust-pptx) -- Rust PPTX reading/writing crate
