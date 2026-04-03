# deckmint-math

LaTeX and MathML to Office Math (OMML) converter for PowerPoint presentations.

This crate is part of [deckmint](https://crates.io/crates/deckmint) and is enabled by default via the `math` feature. It converts LaTeX equations into native, editable OMML that PowerPoint can render and edit.

```rust
use deckmint::Presentation;
use deckmint::objects::text::TextOptionsBuilder;

let mut pres = Presentation::new();
let slide = pres.add_slide();
slide.add_equation(
    r"\int_{0}^{\infty} e^{-x^2} \, dx = \frac{\sqrt{\pi}}{2}",
    TextOptionsBuilder::new().bounds(1.0, 2.0, 8.0, 1.5).build(),
).unwrap();
```

See the [deckmint documentation](https://github.com/rebo/deckmint) for full usage.
