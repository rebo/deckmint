//! First-class SVG embedding via `add_image_svg` / `add_image_svg_from`.
//!
//! PowerPoint renders SVG natively through an `asvg:svgBlip` extension on a
//! picture whose base blip is a raster fallback — both media parts have to
//! be independently registered, real relationships. This example exercises
//! all three chrome tiers (slide, named layout, master) so the offset-by-N
//! rId rebasing gets a real end-to-end run, not just unit coverage.

use deckmint::objects::image::ImageOptionsBuilder;
use deckmint::objects::text::TextOptionsBuilder;
use deckmint::{Presentation, SlideMasterDef};

fn main() {
    let svg = std::fs::read(concat!(env!("CARGO_MANIFEST_DIR"), "/examples/assets/svg_icon.svg"))
        .expect("svg_icon.svg not found — place it in deckmint/examples/assets/");
    let fallback = std::fs::read(concat!(env!("CARGO_MANIFEST_DIR"), "/examples/assets/svg_icon_fallback.png"))
        .expect("svg_icon_fallback.png not found — place it in deckmint/examples/assets/");

    let mut pres = Presentation::new();

    // Master: SVG chrome common to every slide (bottom-left corner badge).
    let mut master = SlideMasterDef::new("SVG Demo Master");
    master.add_image_svg(
        svg.clone(),
        fallback.clone(),
        ImageOptionsBuilder::new().pos(0.2, 6.9).size(0.6, 0.4).build_opts(),
    );
    pres.define_master(master);

    // Named layout: SVG chrome common to one archetype's slides.
    let layout = pres.define_layout("SVG Layout", None, None);
    layout.add_image_svg(
        svg.clone(),
        fallback.clone(),
        ImageOptionsBuilder::new().pos(8.5, 0.2).size(1.2, 0.8).build_opts(),
    );

    // Slide 1: plain add_image_svg, explicit size (aspect intentionally kept).
    let slide1 = pres.add_slide_with_layout("SVG Layout");
    slide1.add_text(
        "add_image_svg — raw bytes",
        TextOptionsBuilder::new().bounds(0.5, 0.3, 7.5, 0.6).font_size(24.0).bold().build(),
    );
    slide1.add_image_svg(
        svg.clone(),
        fallback.clone(),
        ImageOptionsBuilder::new().pos(0.5, 1.2).size(4.0, 2.667).build_opts(),
    );

    // Slide 2: builder variant, auto-dimensioned from the PNG fallback (w only).
    let slide2 = pres.add_slide();
    slide2.add_text(
        "add_image_svg_from — builder, aspect-derived height",
        TextOptionsBuilder::new().bounds(0.5, 0.3, 9.0, 0.6).font_size(22.0).bold().build(),
    );
    slide2
        .add_image_svg_from(
            ImageOptionsBuilder::new()
                .bytes(svg, "svg")
                .svg_fallback(fallback)
                .pos(0.5, 1.2)
                .w(5.0), // height derives from the PNG fallback's intrinsic 3:2 aspect
        )
        .unwrap();

    pres.write_to_file("35_svg_images.pptx").unwrap();
    println!("Wrote 35_svg_images.pptx");
}
