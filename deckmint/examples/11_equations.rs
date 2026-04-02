//! LaTeX equations rendered as native, editable OMML in PowerPoint.
//!
//! Requires the `math` feature (enabled by default).

use deckmint::objects::text::{TextOptionsBuilder, TextRunBuilder};
use deckmint::{AlignH, Presentation};

fn main() {
    let mut pres = Presentation::new();
    pres.title = "Equations Demo".to_string();

    // ── Slide 1: Standalone equations ────────────────────────

    let s = pres.add_slide();
    s.add_text(
        "LaTeX Equations",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.3, 9.0, 0.8)
            .font_size(28.0).bold()
            .build(),
    );

    // Quadratic formula
    s.add_equation(
        r"x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}",
        TextOptionsBuilder::new()
            .bounds(1.0, 1.5, 8.0, 1.2)
            .font_size(24.0)
            .align(AlignH::Center)
            .build(),
    ).unwrap();

    // Euler's identity
    s.add_equation(
        r"e^{i\pi} + 1 = 0",
        TextOptionsBuilder::new()
            .bounds(1.0, 3.0, 8.0, 1.0)
            .font_size(28.0)
            .align(AlignH::Center)
            .build(),
    ).unwrap();

    // Summation
    s.add_equation(
        r"\sum_{n=1}^{\infty} \frac{1}{n^2} = \frac{\pi^2}{6}",
        TextOptionsBuilder::new()
            .bounds(1.0, 4.2, 8.0, 1.2)
            .font_size(24.0)
            .align(AlignH::Center)
            .build(),
    ).unwrap();

    // ── Slide 2: Inline math mixed with text ────────────────

    let s = pres.add_slide();
    s.add_text(
        "Inline Equations",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.3, 9.0, 0.8)
            .font_size(28.0).bold()
            .build(),
    );

    s.add_text_with_math(
        r"The area of a circle is $A = \pi r^2$ where $r$ is the radius.",
        TextOptionsBuilder::new()
            .bounds(0.5, 1.5, 9.0, 1.0)
            .font_size(18.0)
            .build(),
    ).unwrap();

    s.add_text_with_math(
        r"Einstein showed that $E = mc^2$, relating energy and mass.",
        TextOptionsBuilder::new()
            .bounds(0.5, 2.8, 9.0, 1.0)
            .font_size(18.0)
            .build(),
    ).unwrap();

    s.add_text_with_math(
        r"If $a = 5$ and $b = 3$, then $a^2 + b^2 = 34$.",
        TextOptionsBuilder::new()
            .bounds(0.5, 4.0, 9.0, 1.0)
            .font_size(18.0)
            .build(),
    ).unwrap();

    // ── Slide 3: Rich equation runs via TextRunBuilder ──────

    let s = pres.add_slide();
    s.add_text(
        "Equation Runs via Builder",
        TextOptionsBuilder::new()
            .bounds(0.5, 0.3, 9.0, 0.8)
            .font_size(28.0).bold()
            .build(),
    );

    // Mix plain text runs with equation runs in one text box
    s.add_text_runs(
        vec![
            TextRunBuilder::new("The integral ").font_size(18.0).build(),
            TextRunBuilder::equation(r"\int_0^1 x^2 \, dx = \frac{1}{3}").unwrap().build(),
            TextRunBuilder::new(" is a classic result.").font_size(18.0).build(),
        ],
        TextOptionsBuilder::new()
            .bounds(0.5, 1.5, 9.0, 1.0)
            .build(),
    );

    // Matrix example
    s.add_equation(
        r"A = \begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}",
        TextOptionsBuilder::new()
            .bounds(1.0, 3.0, 8.0, 1.5)
            .font_size(24.0)
            .align(AlignH::Center)
            .build(),
    ).unwrap();

    pres.write_to_file("11_equations.pptx").unwrap();
    println!("Wrote 11_equations.pptx");
}
