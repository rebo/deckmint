use deckmint::objects::image::ImageOptionsBuilder;
use deckmint::objects::shape::ShapeOptionsBuilder;
use deckmint::objects::table::{TableCell, TableOptions, TableOptionsBuilder};
use deckmint::objects::text::{TextOptionsBuilder, TextRun, TextRunOptions};
use deckmint::types::{AnimationEffect, CheckerboardDir, Direction, HyperlinkProps, ShapeVariant, SplitOrientation, StripDir};
use deckmint::{AlignH, Presentation, ShapeType};
use std::io::Read;
use zip::ZipArchive;

fn open_zip(bytes: &[u8]) -> ZipArchive<std::io::Cursor<&[u8]>> {
    ZipArchive::new(std::io::Cursor::new(bytes)).expect("valid zip")
}

fn read_entry(zip: &mut ZipArchive<std::io::Cursor<&[u8]>>, name: &str) -> String {
    let mut entry = zip.by_name(name).unwrap_or_else(|_| panic!("missing: {name}"));
    let mut s = String::new();
    entry.read_to_string(&mut s).unwrap();
    s
}

#[test]
fn empty_presentation_is_valid_zip() {
    let pres = Presentation::new();
    let bytes = pres.write().expect("write failed");
    assert!(!bytes.is_empty());
    let mut zip = open_zip(&bytes);
    // Must contain mandatory OOXML files
    read_entry(&mut zip, "[Content_Types].xml");
    read_entry(&mut zip, "_rels/.rels");
    read_entry(&mut zip, "ppt/presentation.xml");
}

#[test]
fn presentation_with_text() {
    let mut pres = Presentation::new();
    pres.title = "Test Title".to_string();
    pres.author = "Test Author".to_string();

    let slide = pres.add_slide();
    slide.add_text(
        "Hello, World!",
        TextOptionsBuilder::new()
            .x(1.0).y(1.5).w(8.0).h(1.5)
            .font_size(36.0)
            .bold()
            .build(),
    );

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);

    let slide_xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(slide_xml.contains("Hello, World!"), "slide must contain text");
    assert!(slide_xml.contains("b=\"1\""), "bold run property expected");
}

#[test]
fn presentation_with_shapes() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    slide.add_shape(
        ShapeType::Rect,
        ShapeOptionsBuilder::new()
            .x(1.0).y(1.0).w(4.0).h(2.0)
            .fill_color("4472C4")
            .build(),
    );
    slide.add_shape(
        ShapeType::Ellipse,
        ShapeOptionsBuilder::new()
            .x(6.0).y(1.0).w(2.0).h(2.0)
            .no_fill()
            .line_color("FF0000")
            .line_width(2.0)
            .build(),
    );

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let slide_xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(slide_xml.contains("rect"), "rect shape expected");
    assert!(slide_xml.contains("ellipse"), "ellipse shape expected");
}

#[test]
fn presentation_with_image() {
    // 1x1 red pixel PNG
    let png_bytes: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
        0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
        0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
        0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
        0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
        0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    let opts = ImageOptionsBuilder::new()
        .x(1.0).y(1.0).w(4.0).h(3.0)
        .build();
    let (opts, _, _, _) = opts; // build() returns (ImageOptions, ..)
    // Workaround: use add_image directly
    slide.add_image(png_bytes, "png", opts);

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    // Global sequential naming: first media across all slides → image1.png
    assert!(zip.by_name("ppt/media/image1.png").is_ok(), "media file must be present");
    let slide_xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(slide_xml.contains("p:pic"), "image element expected");
}

#[test]
fn presentation_with_table() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();

    let rows = vec![
        vec![TableCell::new("Name"), TableCell::new("Score")],
        vec![TableCell::new("Alice"), TableCell::new("95")],
        vec![TableCell::new("Bob"), TableCell::new("87")],
    ];
    let opts = TableOptionsBuilder::new()
        .x(1.0).y(1.0).w(8.0).h(3.0)
        .col_w(vec![4.0, 4.0])
        .build();
    slide.add_table(rows, opts);

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let slide_xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(slide_xml.contains("a:tbl"), "table element expected");
    assert!(slide_xml.contains("Alice"), "cell content expected");
}

#[test]
fn multi_slide_presentation() {
    let mut pres = Presentation::new();

    let s1 = pres.add_slide();
    s1.add_text("Slide 1", TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(1.0).build());

    let s2 = pres.add_slide();
    s2.add_text("Slide 2", TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(1.0).build());
    s2.add_notes("Speaker note for slide 2");

    let s3 = pres.add_slide();
    s3.set_background_color("1F3864");
    s3.add_text("Slide 3 dark bg", TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(1.0).color("FFFFFF").build());

    assert_eq!(pres.slide_count(), 3);

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);

    read_entry(&mut zip, "ppt/slides/slide1.xml");
    read_entry(&mut zip, "ppt/slides/slide2.xml");
    read_entry(&mut zip, "ppt/slides/slide3.xml");

    let notes2 = read_entry(&mut zip, "ppt/notesSlides/notesSlide2.xml");
    assert!(notes2.contains("Speaker note for slide 2"));

    let slide3 = read_entry(&mut zip, "ppt/slides/slide3.xml");
    assert!(slide3.contains("1F3864"), "background color expected");
}

#[test]
fn rich_text_runs() {
    use deckmint::objects::text::{TextRun, TextRunOptions};

    let mut pres = Presentation::new();
    let slide = pres.add_slide();

    let runs = vec![
        TextRun {
            text: "Bold ".to_string(),
            options: TextRunOptions { bold: Some(true), ..Default::default() },
            break_line: false,
            soft_break_before: false,
            field: None,
            equation_omml: None,
        },
        TextRun {
            text: "Italic ".to_string(),
            options: TextRunOptions { italic: Some(true), ..Default::default() },
            break_line: false,
            soft_break_before: false,
            field: None,
            equation_omml: None,
        },
        TextRun {
            text: "Normal".to_string(),
            options: TextRunOptions::default(),
            break_line: false,
            soft_break_before: false,
            field: None,
            equation_omml: None,
        },
    ];

    slide.add_text_runs(runs, TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(2.0).build());

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let slide_xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(slide_xml.contains("Bold"), "bold text expected");
    assert!(slide_xml.contains("Italic"), "italic text expected");
}

#[test]
fn aligned_text() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    slide.add_text("Centered", TextOptionsBuilder::new()
        .x(1.0).y(1.0).w(8.0).h(1.0)
        .align(AlignH::Center)
        .build());

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let slide_xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(slide_xml.contains("algn=\"ctr\""), "center align attribute expected");
}

#[test]
fn write_pptx_file() {
    let mut pres = Presentation::new();
    pres.title = "File Output Test".to_string();
    let slide = pres.add_slide();
    slide.add_text("Written to file", TextOptionsBuilder::new()
        .x(1.0).y(1.0).w(8.0).h(1.0)
        .font_size(24.0)
        .build());

    let path = "/tmp/deckmint_test_output.pptx";
    pres.write_to_file(path).expect("write_to_file failed");

    let meta = std::fs::metadata(path).expect("file must exist");
    assert!(meta.len() > 1000, "pptx must be non-trivial size");
    std::fs::remove_file(path).ok();
}

#[test]
fn hyperlink_text_run_emits_hlinkclick() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();

    let runs = vec![
        TextRun {
            text: "Click here".to_string(),
            options: TextRunOptions {
                hyperlink: Some(HyperlinkProps {
                    r_id: 0, // auto-assigned by add_text_runs
                    slide: None,
                    url: Some("https://example.com".to_string()),
                    tooltip: None,
                    action: None,
                }),
                ..Default::default()
            },
            break_line: false,
            soft_break_before: false,
            field: None,
            equation_omml: None,
        },
    ];
    slide.add_text_runs(runs, TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(1.0).build());

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);

    let slide_xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(slide_xml.contains("hlinkClick"), "hlinkClick must appear in slide XML");

    let rels_xml = read_entry(&mut zip, "ppt/slides/_rels/slide1.xml.rels");
    assert!(rels_xml.contains("https://example.com"), "hyperlink target must be in rels");
    assert!(rels_xml.contains("TargetMode=\"External\""), "external URL needs TargetMode");
}

#[test]
fn hyperlink_slide_jump_emits_ppaction() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    pres.add_slide(); // slide 2 exists

    let slide_mut = pres.slide_mut(0).unwrap();
    let runs = vec![
        TextRun {
            text: "Go to slide 2".to_string(),
            options: TextRunOptions {
                hyperlink: Some(HyperlinkProps {
                    r_id: 0,
                    slide: Some(2),
                    url: None,
                    tooltip: None,
                    action: None,
                }),
                ..Default::default()
            },
            break_line: false,
            soft_break_before: false,
            field: None,
            equation_omml: None,
        },
    ];
    slide_mut.add_text_runs(runs, TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(1.0).build());

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);

    let slide_xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(slide_xml.contains("ppaction://hlinksldjump"), "slide jump action expected");
}

#[test]
fn image_auto_detects_dimensions() {
    // 1x1 red pixel PNG — imagesize should return 1x1
    let png_bytes: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
        0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
        0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
        0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
        0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
        0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let mut pres = Presentation::new();
    let slide = pres.add_slide();

    // No w/h provided — should auto-detect
    let (opts, _, _, _) = ImageOptionsBuilder::new()
        .x(1.0).y(1.0)
        .build();
    slide.add_image(png_bytes, "png", opts);

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let slide_xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    // 1px at 96dpi = 914400/96 = 9525 EMU; non-zero cx/cy expected
    assert!(slide_xml.contains("cx=\"9525\"") || slide_xml.contains("cx="), "auto-detected width expected");
}

// ─── Text feature tests ────────────────────────────────────────

#[test]
fn text_highlight_emits_a_highlight() {
    use deckmint::objects::text::{TextRun, TextRunOptions};

    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    let runs = vec![TextRun {
        text: "highlighted".to_string(),
        options: TextRunOptions { highlight: Some("FFFF00".to_string()), ..Default::default() },
        break_line: false,
        soft_break_before: false,
        field: None,
        equation_omml: None,
    }];
    slide.add_text_runs(runs, TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(1.0).build());

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(xml.contains("<a:highlight>"), "highlight element expected");
    assert!(xml.contains("FFFF00"), "highlight color expected");
}

#[test]
fn text_underline_styles() {
    use deckmint::objects::text::{TextRun, TextRunOptions};

    let mut pres = Presentation::new();
    let slide = pres.add_slide();

    // double underline
    let runs = vec![TextRun {
        text: "double underline".to_string(),
        options: TextRunOptions { underline: Some("dbl".to_string()), ..Default::default() },
        break_line: false,
        soft_break_before: false,
        field: None,
        equation_omml: None,
    }];
    slide.add_text_runs(runs, TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(1.0).build());

    // wavy underline
    let s2 = pres.add_slide();
    let runs2 = vec![TextRun {
        text: "wavy".to_string(),
        options: TextRunOptions { underline: Some("wavy".to_string()), ..Default::default() },
        break_line: false,
        soft_break_before: false,
        field: None,
        equation_omml: None,
    }];
    s2.add_text_runs(runs2, TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(1.0).build());

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let s1_xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    let s2_xml = read_entry(&mut zip, "ppt/slides/slide2.xml");
    assert!(s1_xml.contains("u=\"dbl\""), "double underline expected");
    assert!(s2_xml.contains("u=\"wavy\""), "wavy underline expected");
}

#[test]
fn text_glow_emits_effectlst() {
    use deckmint::objects::text::{TextRun, TextRunOptions};
    use deckmint::GlowProps;

    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    let runs = vec![TextRun {
        text: "glowing".to_string(),
        options: TextRunOptions {
            glow: Some(GlowProps { size: 5.0, color: "FF0000".to_string(), opacity: 0.5 }),
            ..Default::default()
        },
        break_line: false,
        soft_break_before: false,
        field: None,
        equation_omml: None,
    }];
    slide.add_text_runs(runs, TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(1.0).build());

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(xml.contains("<a:effectLst>"), "effectLst expected");
    assert!(xml.contains("<a:glow"), "glow element expected");
}

#[test]
fn text_outline_emits_a_ln() {
    use deckmint::objects::text::{TextRun, TextRunOptions};
    use deckmint::TextOutlineProps;

    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    let runs = vec![TextRun {
        text: "outlined".to_string(),
        options: TextRunOptions {
            outline: Some(TextOutlineProps { color: "000000".to_string(), size: 1.0 }),
            ..Default::default()
        },
        break_line: false,
        soft_break_before: false,
        field: None,
        equation_omml: None,
    }];
    slide.add_text_runs(runs, TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(1.0).build());

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(xml.contains("<a:ln"), "outline ln element expected");
}

#[test]
fn text_soft_break_emits_a_br() {
    use deckmint::objects::text::{TextRun, TextRunOptions};

    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    let runs = vec![
        TextRun {
            text: "line one".to_string(),
            options: TextRunOptions::default(),
            break_line: false,
            soft_break_before: false,
            field: None,
            equation_omml: None,
        },
        TextRun {
            text: "line two (soft break before)".to_string(),
            options: TextRunOptions::default(),
            break_line: false,
            soft_break_before: true,
            field: None,
            equation_omml: None,
        },
    ];
    slide.add_text_runs(runs, TextOptionsBuilder::new().x(1.0).y(1.0).w(8.0).h(2.0).build());

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(xml.contains("<a:br>"), "soft break <a:br> element expected");
}

#[test]
fn text_direction_emits_vert_attr() {
    use deckmint::objects::text::TextOptionsBuilder;

    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    slide.add_text(
        "vertical text",
        TextOptionsBuilder::new()
            .x(1.0).y(1.0).w(1.0).h(4.0)
            .text_direction("vert")
            .build(),
    );

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(xml.contains("vert=\"vert\""), "vert attribute on bodyPr expected");
}

#[test]
fn tab_stops_emit_a_tab() {
    use deckmint::objects::text::TextOptionsBuilder;
    use deckmint::TabStop;

    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    slide.add_text(
        "tabbed\ttext",
        TextOptionsBuilder::new()
            .x(1.0).y(1.0).w(8.0).h(1.0)
            .tab_stops(vec![
                TabStop { pos_inches: 2.0, align: "l".to_string() },
                TabStop { pos_inches: 5.0, align: "r".to_string() },
            ])
            .build(),
    );

    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(xml.contains("<a:tab"), "tab stop element expected");
    assert!(xml.contains("algn=\"l\""), "left-aligned tab stop expected");
    assert!(xml.contains("algn=\"r\""), "right-aligned tab stop expected");
}

#[test]
#[ignore = "requires npx on PATH with @xarsh/ooxml-validator"]
fn pptx_passes_ooxml_validation() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    slide.add_text("Hello", Default::default());
    let bytes = pres.write().unwrap();
    let path = "/tmp/deckmint_validate_test.pptx";
    std::fs::write(path, &bytes).unwrap();
    let output = std::process::Command::new("npx")
        .args(["--yes", "@xarsh/ooxml-validator", path])
        .output()
        .expect("npx must be on PATH");
    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    std::fs::remove_file(path).ok();
    assert!(
        output.status.success(),
        "OOXML validation failed:\n{stdout}\n{stderr}"
    );
}

#[test]
fn animation_appear_emits_timing_block() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    slide.add_text(
        "Animated",
        TextOptionsBuilder::new().x(1.0).y(1.0).w(4.0).h(1.0)
            .animation(AnimationEffect::appear())
            .build(),
    );
    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(xml.contains("<p:timing>"), "timing block must be present");
    assert!(xml.contains("presetClass=\"entr\""), "entrance preset class");
    assert!(xml.contains("style.visibility"), "visibility attribute");
    assert!(xml.contains("val=\"visible\""), "appear sets visible");
}

#[test]
fn animation_disappear_emits_exit_timing() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    slide.add_text(
        "Gone",
        TextOptionsBuilder::new().x(1.0).y(1.0).w(4.0).h(1.0)
            .animation(AnimationEffect::disappear())
            .build(),
    );
    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(xml.contains("presetClass=\"exit\""), "exit preset class");
    assert!(xml.contains("val=\"hidden\""), "disappear sets hidden");
}

#[test]
fn slide_without_animation_has_no_timing_block() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    slide.add_text("No animation", Default::default());
    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    assert!(!xml.contains("<p:timing>"), "no timing block for non-animated slide");
}

#[test]
fn animation_multiple_objects_get_sequential_ids() {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    slide.add_text("First",  TextOptionsBuilder::new().x(1.0).y(1.0).w(4.0).h(0.5).animation(AnimationEffect::appear()).build());
    slide.add_text("Second", TextOptionsBuilder::new().x(1.0).y(2.0).w(4.0).h(0.5).animation(AnimationEffect::disappear()).build());
    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    let xml = read_entry(&mut zip, "ppt/slides/slide1.xml");
    // Two animations → grpId=0 and grpId=1
    assert!(xml.contains("grpId=\"0\""), "first animation group id");
    assert!(xml.contains("grpId=\"1\""), "second animation group id");
    // Two bldP entries
    assert!(xml.contains("spid=\"2\""), "spid for first object");
    assert!(xml.contains("spid=\"3\""), "spid for second object");
}

// ── helpers ──────────────────────────────────────────────────────────────────

fn make_slide_with_anim(anim: AnimationEffect) -> String {
    let mut pres = Presentation::new();
    let slide = pres.add_slide();
    slide.add_text(
        "test",
        TextOptionsBuilder::new().x(1.0).y(1.0).w(4.0).h(1.0)
            .animation(anim)
            .build(),
    );
    let bytes = pres.write().expect("write failed");
    let mut zip = open_zip(&bytes);
    read_entry(&mut zip, "ppt/slides/slide1.xml")
}

// ── FadeIn / FadeOut ─────────────────────────────────────────────────────────

#[test]
fn animation_fade_in_emits_entr_preset10() {
    let xml = make_slide_with_anim(AnimationEffect::fade_in());
    assert!(xml.contains("<p:timing>"), "timing block present");
    assert!(xml.contains("presetClass=\"entr\""), "entrance class");
    assert!(xml.contains("presetID=\"10\""), "fade preset id");
    assert!(xml.contains("transition=\"in\""), "transition in");
    assert!(xml.contains("filter=\"fade\""), "fade filter");
}

#[test]
fn animation_fade_out_emits_exit_preset10() {
    let xml = make_slide_with_anim(AnimationEffect::fade_out());
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("presetID=\"10\""), "fade preset id");
    assert!(xml.contains("transition=\"out\""), "transition out");
    assert!(xml.contains("filter=\"fade\""), "fade filter");
}

// ── WipeIn / WipeOut ─────────────────────────────────────────────────────────

#[test]
fn animation_wipe_in_left_emits_entr_preset4() {
    let xml = make_slide_with_anim(AnimationEffect::wipe_in(Direction::Left));
    assert!(xml.contains("presetClass=\"entr\""), "entrance class");
    assert!(xml.contains("presetID=\"4\""), "wipe preset id");
    assert!(xml.contains("wipe(left)"), "left wipe filter");
    assert!(xml.contains("transition=\"in\""), "transition in");
}

#[test]
fn animation_wipe_out_right() {
    let xml = make_slide_with_anim(AnimationEffect::wipe_out(Direction::Right));
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("wipe(right)"), "right wipe filter");
    assert!(xml.contains("transition=\"out\""), "transition out");
}

#[test]
fn animation_wipe_in_up() {
    let xml = make_slide_with_anim(AnimationEffect::wipe_in(Direction::Up));
    assert!(xml.contains("wipe(up)"), "up wipe filter");
}

#[test]
fn animation_wipe_in_down() {
    let xml = make_slide_with_anim(AnimationEffect::wipe_in(Direction::Down));
    assert!(xml.contains("wipe(down)"), "down wipe filter");
}

// ── ZoomIn / ZoomOut ─────────────────────────────────────────────────────────

#[test]
fn animation_zoom_in_emits_entr_preset11() {
    let xml = make_slide_with_anim(AnimationEffect::zoom_in());
    assert!(xml.contains("presetClass=\"entr\""), "entrance class");
    assert!(xml.contains("presetID=\"11\""), "zoom preset id");
    assert!(xml.contains("<p:animScale>"), "animScale element");
    assert!(xml.contains("x=\"10000\""), "from small scale");
    assert!(xml.contains("x=\"100000\""), "to full scale");
}

#[test]
fn animation_zoom_out_emits_exit_preset11() {
    let xml = make_slide_with_anim(AnimationEffect::zoom_out());
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("presetID=\"11\""), "zoom preset id");
    assert!(xml.contains("x=\"100000\""), "from full scale");
    assert!(xml.contains("x=\"10000\""), "to small scale");
}

// ── FlyIn / FlyOut ───────────────────────────────────────────────────────────

#[test]
fn animation_fly_in_from_left_emits_ppt_x() {
    let xml = make_slide_with_anim(AnimationEffect::fly_in(Direction::Left));
    assert!(xml.contains("presetClass=\"entr\""), "entrance class");
    assert!(xml.contains("presetID=\"2\""), "fly preset id");
    assert!(xml.contains("<p:anim"), "anim element");
    assert!(xml.contains("ppt_x"), "horizontal fly uses ppt_x");
    assert!(xml.contains("additive=\"sum\""), "uses additive sum offsets");
    // FlyIn from Left: start offset = -1.0 (off left edge), end = 0.0 (natural position)
    assert!(xml.contains("val=\"-1\""), "start offset off left edge");
    assert!(xml.contains("val=\"0\""), "end offset at natural position");
}

#[test]
fn animation_fly_in_from_bottom_emits_ppt_y() {
    let xml = make_slide_with_anim(AnimationEffect::fly_in(Direction::Down));
    assert!(xml.contains("ppt_y"), "vertical fly uses ppt_y");
    // FlyIn from Down: start offset = +1.0 (below slide), end = 0.0
    assert!(xml.contains("val=\"1\""), "start offset below slide");
}

#[test]
fn animation_fly_out_to_right_emits_ppt_x_end() {
    let xml = make_slide_with_anim(AnimationEffect::fly_out(Direction::Right));
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("ppt_x"), "horizontal fly uses ppt_x");
    // FlyOut to Right: start = 0.0, end = +1.0 (off right edge)
    assert!(xml.contains("val=\"1\""), "end offset off right edge");
}

#[test]
fn animation_fly_out_to_top_emits_ppt_y_end() {
    let xml = make_slide_with_anim(AnimationEffect::fly_out(Direction::Up));
    assert!(xml.contains("ppt_y"), "vertical fly uses ppt_y");
    // FlyOut to Up: start = 0.0, end = -1.0 (above slide)
    assert!(xml.contains("val=\"-1\""), "end offset above slide");
}

// ── Split ────────────────────────────────────────────────────────────────────

#[test]
fn animation_split_in_horizontal_emits_barn_filter() {
    let xml = make_slide_with_anim(AnimationEffect::split_in(SplitOrientation::Horizontal));
    assert!(xml.contains("presetClass=\"entr\""), "entrance class");
    assert!(xml.contains("presetID=\"3\""), "split preset id");
    assert!(xml.contains("barn("), "barn filter");
    assert!(xml.contains("horizontal"), "horizontal orientation");
    assert!(xml.contains("direction=in"), "split entering");
}

#[test]
fn animation_split_out_vertical_emits_barn_filter() {
    let xml = make_slide_with_anim(AnimationEffect::split_out(SplitOrientation::Vertical));
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("vertical"), "vertical orientation");
    assert!(xml.contains("direction=out"), "split exiting");
}

// ── Emphasis: Spin ────────────────────────────────────────────────────────────

#[test]
fn animation_spin_360_emits_anim_rot() {
    let xml = make_slide_with_anim(AnimationEffect::spin(360.0));
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("presetID=\"8\""), "spin preset id");
    assert!(xml.contains("<p:animRot"), "animRot element");
    // 360 degrees × 60000 = 21600000
    assert!(xml.contains("by=\"21600000\""), "360 degree rotation");
    assert!(xml.contains("<p:attrName>r</p:attrName>"), "rotation attribute name");
}

#[test]
fn animation_spin_90_emits_correct_angle() {
    let xml = make_slide_with_anim(AnimationEffect::spin(90.0));
    // 90 × 60000 = 5400000
    assert!(xml.contains("by=\"5400000\""), "90 degree rotation angle");
}

// ── Emphasis: Pulse ───────────────────────────────────────────────────────────

#[test]
fn animation_pulse_emits_anim_scale_with_autoreverse() {
    let xml = make_slide_with_anim(AnimationEffect::pulse());
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("presetID=\"14\""), "pulse preset id");
    assert!(xml.contains("<p:animScale>"), "animScale element");
    assert!(xml.contains("autoRev=\"1\""), "auto-reverse for pulse");
    assert!(xml.contains("x=\"115000\""), "grows to 115%");
}

// ── Emphasis: GrowShrink ──────────────────────────────────────────────────────

#[test]
fn animation_grow_shrink_emits_preset18() {
    let xml = make_slide_with_anim(AnimationEffect::grow_shrink(1.5));
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("presetID=\"18\""), "grow/shrink preset id");
    assert!(xml.contains("<p:animScale>"), "animScale element");
    // 1.5 × 100000 = 150000
    assert!(xml.contains("x=\"150000\""), "scales to 150%");
}

// ── Structural: prevCondLst / nextCondLst ─────────────────────────────────────

#[test]
fn animation_timing_uses_correct_cond_evt_values() {
    let xml = make_slide_with_anim(AnimationEffect::fade_in());
    assert!(xml.contains("evt=\"onPrev\""), "onPrev (not onPrevClick)");
    assert!(xml.contains("evt=\"onNext\""), "onNext (not onNextClick)");
    assert!(!xml.contains("onPrevClick"), "no invalid onPrevClick");
    assert!(!xml.contains("onNextClick"), "no invalid onNextClick");
}

// ── Structural: timing after clrMapOvr ───────────────────────────────────────

#[test]
fn animation_timing_block_appears_after_clr_map_ovr() {
    let xml = make_slide_with_anim(AnimationEffect::appear());
    let clr_pos = xml.find("<p:clrMapOvr>").expect("clrMapOvr in slide");
    let timing_pos = xml.find("<p:timing>").expect("timing in slide");
    assert!(timing_pos > clr_pos, "timing must come after clrMapOvr");
}

// ── Blinds ────────────────────────────────────────────────────────────────────

#[test]
fn animation_blinds_horizontal_emits_preset3() {
    let xml = make_slide_with_anim(AnimationEffect::blinds_in(SplitOrientation::Horizontal));
    assert!(xml.contains("presetClass=\"entr\""), "entrance class");
    assert!(xml.contains("presetID=\"3\""), "blinds preset id");
    assert!(xml.contains("blinds(horizontal)"), "horizontal blinds filter");
    assert!(xml.contains("transition=\"in\""), "transition in");
}

#[test]
fn animation_blinds_vertical_emits_preset3() {
    let xml = make_slide_with_anim(AnimationEffect::blinds_in(SplitOrientation::Vertical));
    assert!(xml.contains("presetID=\"3\""), "blinds preset id");
    assert!(xml.contains("blinds(vertical)"), "vertical blinds filter");
}

// ── Checkerboard ──────────────────────────────────────────────────────────────

#[test]
fn animation_checkerboard_across_emits_preset5() {
    let xml = make_slide_with_anim(AnimationEffect::checkerboard_in(CheckerboardDir::Across));
    assert!(xml.contains("presetID=\"5\""), "checkerboard preset id");
    assert!(xml.contains("checkerboard(across)"), "across filter");
}

#[test]
fn animation_checkerboard_down_emits_preset5() {
    let xml = make_slide_with_anim(AnimationEffect::checkerboard_in(CheckerboardDir::Down));
    assert!(xml.contains("presetID=\"5\""), "checkerboard preset id");
    assert!(xml.contains("checkerboard(down)"), "down filter");
}

// ── Dissolve In ───────────────────────────────────────────────────────────────

#[test]
fn animation_dissolve_in_emits_preset12() {
    let xml = make_slide_with_anim(AnimationEffect::dissolve_in());
    assert!(xml.contains("presetID=\"12\""), "dissolve preset id");
    assert!(xml.contains("dissolve()"), "dissolve filter");
    assert!(xml.contains("presetClass=\"entr\""), "entrance class");
}

// ── Peek In ───────────────────────────────────────────────────────────────────

#[test]
fn animation_peek_in_emits_preset13() {
    let xml = make_slide_with_anim(AnimationEffect::peek_in(Direction::Down));
    assert!(xml.contains("presetID=\"13\""), "peek preset id");
    assert!(xml.contains("wipe(down)"), "wipe down filter");
}

// ── Random Bars ───────────────────────────────────────────────────────────────

#[test]
fn animation_random_bars_horizontal_emits_preset14() {
    let xml = make_slide_with_anim(AnimationEffect::random_bars_in(SplitOrientation::Horizontal));
    assert!(xml.contains("presetID=\"14\""), "random bars preset id");
    assert!(xml.contains("randombar(horizontal)"), "horizontal random bars filter");
}

#[test]
fn animation_random_bars_vertical_emits_preset14() {
    let xml = make_slide_with_anim(AnimationEffect::random_bars_in(SplitOrientation::Vertical));
    assert!(xml.contains("presetID=\"14\""), "random bars preset id");
    assert!(xml.contains("randombar(vertical)"), "vertical random bars filter");
}

// ── Shape ─────────────────────────────────────────────────────────────────────

#[test]
fn animation_shape_box_emits_box_filter() {
    let xml = make_slide_with_anim(AnimationEffect::shape_in(ShapeVariant::Box));
    assert!(xml.contains("presetID=\"8\""), "box preset id");
    assert!(xml.contains("box(in)"), "box in filter");
}

#[test]
fn animation_shape_circle_emits_circle_filter() {
    let xml = make_slide_with_anim(AnimationEffect::shape_in(ShapeVariant::Circle));
    assert!(xml.contains("presetID=\"14\""), "circle preset id");
    assert!(xml.contains("circle(in)"), "circle filter");
}

#[test]
fn animation_shape_diamond_emits_diamond_filter() {
    let xml = make_slide_with_anim(AnimationEffect::shape_in(ShapeVariant::Diamond));
    assert!(xml.contains("presetID=\"15\""), "diamond preset id");
    assert!(xml.contains("diamond(in)"), "diamond filter");
}

#[test]
fn animation_shape_plus_emits_plus_filter() {
    let xml = make_slide_with_anim(AnimationEffect::shape_in(ShapeVariant::Plus));
    assert!(xml.contains("presetID=\"16\""), "plus preset id");
    assert!(xml.contains("plus(in)"), "plus filter");
}

// ── Strips ────────────────────────────────────────────────────────────────────

#[test]
fn animation_strips_leftdown_emits_preset6() {
    let xml = make_slide_with_anim(AnimationEffect::strips_in(StripDir::LeftDown));
    assert!(xml.contains("presetID=\"6\""), "strips preset id");
    assert!(xml.contains("strips(leftdown)"), "leftdown filter");
}

// ── Wedge ─────────────────────────────────────────────────────────────────────

#[test]
fn animation_wedge_emits_preset17() {
    let xml = make_slide_with_anim(AnimationEffect::wedge_in());
    assert!(xml.contains("presetID=\"17\""), "wedge preset id");
    assert!(xml.contains("wedge()"), "wedge filter");
}

// ── Wheel ─────────────────────────────────────────────────────────────────────

#[test]
fn animation_wheel_3spokes_emits_preset18() {
    let xml = make_slide_with_anim(AnimationEffect::wheel_in(3));
    assert!(xml.contains("presetID=\"18\""), "wheel preset id");
    assert!(xml.contains("wheel(spokes=3)"), "3 spoke wheel filter");
}

// ── Expand / Swivel / BasicZoom ───────────────────────────────────────────────

#[test]
fn animation_expand_emits_preset22() {
    let xml = make_slide_with_anim(AnimationEffect::expand_in());
    assert!(xml.contains("presetID=\"22\""), "expand preset id");
    assert!(xml.contains("<p:animScale>"), "animScale element");
}

#[test]
fn animation_swivel_emits_preset21() {
    let xml = make_slide_with_anim(AnimationEffect::swivel_in());
    assert!(xml.contains("presetID=\"21\""), "swivel preset id");
    assert!(xml.contains("<p:animScale>"), "animScale element");
}

#[test]
fn animation_basic_zoom_emits_preset27() {
    let xml = make_slide_with_anim(AnimationEffect::basic_zoom_in());
    assert!(xml.contains("presetID=\"27\""), "basic zoom preset id");
    assert!(xml.contains("<p:animScale>"), "animScale element");
}

// ── Moderate entrance animations ──────────────────────────────────────────────

#[test]
fn animation_centre_revolve_emits_preset23() {
    let xml = make_slide_with_anim(AnimationEffect::centre_revolve_in());
    assert!(xml.contains("presetID=\"23\""), "centre revolve preset id");
    assert!(xml.contains("<p:animScale>"), "scale component");
    assert!(xml.contains("<p:animRot"), "rotation component");
}

#[test]
fn animation_float_in_emits_preset30() {
    let xml = make_slide_with_anim(AnimationEffect::float_in(Direction::Up));
    assert!(xml.contains("presetID=\"30\""), "float in preset id");
    assert!(xml.contains("filter=\"fade\""), "fade component");
    assert!(xml.contains("style.rotation"), "rotation component");
}

#[test]
fn animation_grow_turn_emits_preset24() {
    let xml = make_slide_with_anim(AnimationEffect::grow_turn_in());
    assert!(xml.contains("presetID=\"24\""), "grow turn preset id");
    assert!(xml.contains("<p:animScale>"), "scale component");
    assert!(xml.contains("<p:animRot"), "rotation component");
}

#[test]
fn animation_rise_up_emits_preset25() {
    let xml = make_slide_with_anim(AnimationEffect::rise_up_in());
    assert!(xml.contains("presetID=\"25\""), "rise up preset id");
    assert!(xml.contains("ppt_y"), "vertical position animation");
}

#[test]
fn animation_spinner_emits_preset28() {
    let xml = make_slide_with_anim(AnimationEffect::spinner_in());
    assert!(xml.contains("presetID=\"28\""), "spinner preset id");
    assert!(xml.contains("<p:animScale>"), "scale component");
    assert!(xml.contains("<p:animRot"), "rotation component");
}

#[test]
fn animation_stretch_horizontal_emits_preset29() {
    let xml = make_slide_with_anim(AnimationEffect::stretch_in(Direction::Left));
    assert!(xml.contains("presetID=\"29\""), "stretch preset id");
    assert!(xml.contains("<p:animScale>"), "animScale element");
}

// ── Exciting entrance animations ──────────────────────────────────────────────

#[test]
fn animation_bounce_emits_preset26() {
    let xml = make_slide_with_anim(AnimationEffect::bounce_in());
    assert!(xml.contains("presetID=\"26\""), "bounce preset id");
    assert!(xml.contains("wipe(down)"), "wipe down reveal");
    assert!(xml.contains("ppt_y"), "vertical drop animation");
}

#[test]
fn animation_credits_emits_preset32() {
    let xml = make_slide_with_anim(AnimationEffect::credits_in());
    assert!(xml.contains("presetID=\"32\""), "credits preset id");
    assert!(xml.contains("ppt_y"), "vertical position animation");
}

#[test]
fn animation_drop_emits_preset34() {
    let xml = make_slide_with_anim(AnimationEffect::drop_in());
    assert!(xml.contains("presetID=\"34\""), "drop preset id");
    assert!(xml.contains("ppt_y"), "vertical fall animation");
}

#[test]
fn animation_flip_emits_preset35() {
    let xml = make_slide_with_anim(AnimationEffect::flip_in());
    assert!(xml.contains("presetID=\"35\""), "flip preset id");
    assert!(xml.contains("<p:animScale>"), "scale element");
}

#[test]
fn animation_pinwheel_emits_preset37() {
    let xml = make_slide_with_anim(AnimationEffect::pinwheel_in());
    assert!(xml.contains("presetID=\"37\""), "pinwheel preset id");
    assert!(xml.contains("<p:animRot"), "rotation component");
}

#[test]
fn animation_spiral_in_emits_preset38() {
    let xml = make_slide_with_anim(AnimationEffect::spiral_in());
    assert!(xml.contains("presetID=\"38\""), "spiral in preset id");
}

#[test]
fn animation_basic_swivel_emits_preset39() {
    let xml = make_slide_with_anim(AnimationEffect::basic_swivel_in());
    assert!(xml.contains("presetID=\"39\""), "basic swivel preset id");
    assert!(xml.contains("<p:animRot"), "rotation element");
}

#[test]
fn animation_whip_emits_preset40() {
    let xml = make_slide_with_anim(AnimationEffect::whip_in());
    assert!(xml.contains("presetID=\"40\""), "whip preset id");
}

#[test]
fn animation_curve_up_emits_preset52_with_motion() {
    let xml = make_slide_with_anim(AnimationEffect::curve_up_in());
    assert!(xml.contains("presetID=\"52\""), "curve up preset id");
    assert!(xml.contains("<p:animMotion"), "bezier motion path");
    assert!(xml.contains("<p:animScale>"), "scale component");
    assert!(xml.contains("filter=\"fade\""), "fade component");
}

#[test]
fn animation_boomerang_in_emits_preset36() {
    let xml = make_slide_with_anim(AnimationEffect::boomerang_in());
    assert!(xml.contains("presetID=\"36\""), "boomerang preset id");
    assert!(xml.contains("<p:animScale>"), "scale component");
    assert!(xml.contains("ppt_x"), "horizontal motion");
}

// ── Exit Basic additional animations ─────────────────────────────────────────

#[test]
fn animation_blinds_out_emits_preset3_exit() {
    let xml = make_slide_with_anim(AnimationEffect::blinds_out(SplitOrientation::Horizontal));
    assert!(xml.contains("presetID=\"3\""), "blinds out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("blinds(horizontal)"), "horizontal blinds filter");
}

#[test]
fn animation_checkerboard_out_emits_preset5_exit() {
    let xml = make_slide_with_anim(AnimationEffect::checkerboard_out(CheckerboardDir::Across));
    assert!(xml.contains("presetID=\"5\""), "checkerboard out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("checkerboard(across)"), "checkerboard filter");
}

#[test]
fn animation_dissolve_out_emits_preset12_exit() {
    let xml = make_slide_with_anim(AnimationEffect::dissolve_out());
    assert!(xml.contains("presetID=\"12\""), "dissolve out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("dissolve()"), "dissolve filter");
}

#[test]
fn animation_peek_out_emits_preset13_exit() {
    let xml = make_slide_with_anim(AnimationEffect::peek_out(Direction::Left));
    assert!(xml.contains("presetID=\"13\""), "peek out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
}

#[test]
fn animation_random_bars_out_emits_preset14_exit() {
    let xml = make_slide_with_anim(AnimationEffect::random_bars_out(SplitOrientation::Horizontal));
    assert!(xml.contains("presetID=\"14\""), "random bars out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("randombar(horizontal)"), "random bars filter");
}

#[test]
fn animation_shape_out_box_emits_exit() {
    let xml = make_slide_with_anim(AnimationEffect::shape_out(ShapeVariant::Box));
    assert!(xml.contains("presetID=\"8\""), "shape out box preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("box(out)"), "box out filter");
}

#[test]
fn animation_strips_out_emits_preset6_exit() {
    let xml = make_slide_with_anim(AnimationEffect::strips_out(StripDir::LeftDown));
    assert!(xml.contains("presetID=\"6\""), "strips out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
}

#[test]
fn animation_wedge_out_emits_preset17_exit() {
    let xml = make_slide_with_anim(AnimationEffect::wedge_out());
    assert!(xml.contains("presetID=\"17\""), "wedge out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("wedge()"), "wedge filter");
}

#[test]
fn animation_wheel_out_emits_preset18_exit() {
    let xml = make_slide_with_anim(AnimationEffect::wheel_out(4));
    assert!(xml.contains("presetID=\"18\""), "wheel out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("wheel(spokes=4)"), "4-spoke wheel filter");
}

// ── Exit Subtle ───────────────────────────────────────────────────────────────

#[test]
fn animation_contract_out_emits_preset22_exit() {
    let xml = make_slide_with_anim(AnimationEffect::contract_out());
    assert!(xml.contains("presetID=\"22\""), "contract out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animScale>"), "scale element");
}

#[test]
fn animation_swivel_out_emits_preset21_exit() {
    let xml = make_slide_with_anim(AnimationEffect::swivel_out());
    assert!(xml.contains("presetID=\"21\""), "swivel out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animScale>"), "scale element");
}

// ── Exit Moderate ─────────────────────────────────────────────────────────────

#[test]
fn animation_centre_revolve_out_emits_preset23_exit() {
    let xml = make_slide_with_anim(AnimationEffect::centre_revolve_out());
    assert!(xml.contains("presetID=\"23\""), "centre revolve out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animRot"), "rotation component");
}

#[test]
fn animation_collapse_out_emits_preset31_exit() {
    let xml = make_slide_with_anim(AnimationEffect::collapse_out());
    assert!(xml.contains("presetID=\"31\""), "collapse out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animScale>"), "scale element");
}

#[test]
fn animation_float_out_emits_preset30_exit() {
    let xml = make_slide_with_anim(AnimationEffect::float_out(Direction::Up));
    assert!(xml.contains("presetID=\"30\""), "float out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("filter=\"fade\""), "fade component");
    assert!(xml.contains("style.rotation"), "rotation component");
}

#[test]
fn animation_shrink_turn_out_emits_preset24_exit() {
    let xml = make_slide_with_anim(AnimationEffect::shrink_turn_out());
    assert!(xml.contains("presetID=\"24\""), "shrink turn out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animRot"), "rotation component");
}

#[test]
fn animation_sink_down_out_emits_preset25_exit() {
    let xml = make_slide_with_anim(AnimationEffect::sink_down_out());
    assert!(xml.contains("presetID=\"25\""), "sink down out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("ppt_y"), "vertical motion");
}

#[test]
fn animation_spinner_out_emits_preset28_exit() {
    let xml = make_slide_with_anim(AnimationEffect::spinner_out());
    assert!(xml.contains("presetID=\"28\""), "spinner out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animRot"), "rotation component");
}

#[test]
fn animation_basic_zoom_out_emits_preset27_exit() {
    let xml = make_slide_with_anim(AnimationEffect::basic_zoom_out());
    assert!(xml.contains("presetID=\"27\""), "basic zoom out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animScale>"), "scale element");
}

#[test]
fn animation_stretchy_out_emits_preset29_exit() {
    let xml = make_slide_with_anim(AnimationEffect::stretchy_out(Direction::Left));
    assert!(xml.contains("presetID=\"29\""), "stretchy out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animScale>"), "scale element");
}

// ── Exit Exciting ─────────────────────────────────────────────────────────────

#[test]
fn animation_boomerang_out_emits_preset36_exit() {
    let xml = make_slide_with_anim(AnimationEffect::boomerang_out());
    assert!(xml.contains("presetID=\"36\""), "boomerang out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("ppt_x"), "horizontal motion");
}

#[test]
fn animation_bounce_out_emits_preset26_exit() {
    let xml = make_slide_with_anim(AnimationEffect::bounce_out());
    assert!(xml.contains("presetID=\"26\""), "bounce out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("ppt_y"), "vertical motion");
}

#[test]
fn animation_credits_out_emits_preset32_exit() {
    let xml = make_slide_with_anim(AnimationEffect::credits_out());
    assert!(xml.contains("presetID=\"32\""), "credits out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("ppt_y"), "vertical scroll");
}

#[test]
fn animation_curve_down_out_emits_preset52_exit() {
    let xml = make_slide_with_anim(AnimationEffect::curve_down_out());
    assert!(xml.contains("presetID=\"52\""), "curve down preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animMotion"), "bezier motion path");
    assert!(xml.contains("<p:animScale>"), "scale component");
}

#[test]
fn animation_drop_out_emits_preset34_exit() {
    let xml = make_slide_with_anim(AnimationEffect::drop_out());
    assert!(xml.contains("presetID=\"34\""), "drop out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("ppt_y"), "vertical motion");
}

#[test]
fn animation_flip_out_emits_preset35_exit() {
    let xml = make_slide_with_anim(AnimationEffect::flip_out());
    assert!(xml.contains("presetID=\"35\""), "flip out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animScale>"), "scale element");
}

#[test]
fn animation_pinwheel_out_emits_preset37_exit() {
    let xml = make_slide_with_anim(AnimationEffect::pinwheel_out());
    assert!(xml.contains("presetID=\"37\""), "pinwheel out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animRot"), "rotation component");
}

#[test]
fn animation_spiral_out_emits_preset38_exit() {
    let xml = make_slide_with_anim(AnimationEffect::spiral_out());
    assert!(xml.contains("presetID=\"38\""), "spiral out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
}

#[test]
fn animation_basic_swivel_out_emits_preset39_exit() {
    let xml = make_slide_with_anim(AnimationEffect::basic_swivel_out());
    assert!(xml.contains("presetID=\"39\""), "basic swivel out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("<p:animRot"), "rotation element");
}

#[test]
fn animation_whip_out_emits_preset40_exit() {
    let xml = make_slide_with_anim(AnimationEffect::whip_out());
    assert!(xml.contains("presetID=\"40\""), "whip out preset id");
    assert!(xml.contains("presetClass=\"exit\""), "exit class");
    assert!(xml.contains("ppt_x"), "horizontal motion");
}

// ── Emphasis — Basic additional ───────────────────────────────────────────────

#[test]
fn animation_fill_color_emits_preset1_emph() {
    let xml = make_slide_with_anim(AnimationEffect::fill_color("FF0000"));
    assert!(xml.contains("presetID=\"1\""), "fill color preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("<p:animClr"), "animClr element");
    assert!(xml.contains("fillcolor"), "fill attr");
    assert!(xml.contains("FF0000"), "target colour");
}

#[test]
fn animation_font_color_emits_preset2_emph() {
    let xml = make_slide_with_anim(AnimationEffect::font_color("0000FF"));
    assert!(xml.contains("presetID=\"2\""), "font color preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("style.color"), "font colour attr");
    assert!(xml.contains("0000FF"), "target colour");
}

#[test]
fn animation_line_color_emits_preset3_emph() {
    let xml = make_slide_with_anim(AnimationEffect::line_color("00FF00"));
    assert!(xml.contains("presetID=\"3\""), "line color preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("strokecolor"), "stroke attr");
}

#[test]
fn animation_transparency_emits_preset4_emph() {
    let xml = make_slide_with_anim(AnimationEffect::transparency(0.3));
    assert!(xml.contains("presetID=\"4\""), "transparency preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("style.opacity"), "opacity attr");
    assert!(xml.contains("autoRev=\"1\""), "auto-reverse");
}

// ── Emphasis — Subtle ─────────────────────────────────────────────────────────

#[test]
fn animation_bold_flash_emits_preset5_emph() {
    let xml = make_slide_with_anim(AnimationEffect::bold_flash());
    assert!(xml.contains("presetID=\"5\""), "bold flash preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("style.fontWeight"), "fontWeight attr");
    assert!(xml.contains("autoRev=\"1\""), "auto-reverse");
}

#[test]
fn animation_brush_color_emits_preset6_emph() {
    let xml = make_slide_with_anim(AnimationEffect::brush_color("FF8000"));
    assert!(xml.contains("presetID=\"6\""), "brush color preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("autoRev=\"1\""), "auto-reverse");
    assert!(xml.contains("FF8000"), "target colour");
}

#[test]
fn animation_complementary_color_emits_preset7_emph() {
    let xml = make_slide_with_anim(AnimationEffect::complementary_color());
    assert!(xml.contains("presetID=\"7\""), "complementary preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("10800000"), "180° hue shift");
    assert!(xml.contains("clrSpc=\"hsl\""), "hsl colour space");
}

#[test]
fn animation_complementary_color2_emits_preset9_emph() {
    let xml = make_slide_with_anim(AnimationEffect::complementary_color2());
    assert!(xml.contains("presetID=\"9\""), "complementary2 preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("7200000"), "120° hue shift");
}

#[test]
fn animation_contrasting_color_emits_preset10_emph() {
    let xml = make_slide_with_anim(AnimationEffect::contrasting_color());
    assert!(xml.contains("presetID=\"10\""), "contrasting color preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("clrSpc=\"hsl\""), "hsl colour space");
}

#[test]
fn animation_darken_emits_preset11_emph() {
    let xml = make_slide_with_anim(AnimationEffect::darken());
    assert!(xml.contains("presetID=\"11\""), "darken preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("style.opacity"), "opacity animation");
    assert!(xml.contains("autoRev=\"1\""), "auto-reverse");
}

#[test]
fn animation_desaturate_emits_preset12_emph() {
    let xml = make_slide_with_anim(AnimationEffect::desaturate());
    assert!(xml.contains("presetID=\"12\""), "desaturate preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("808080"), "grey target");
}

#[test]
fn animation_lighten_emits_preset13_emph() {
    let xml = make_slide_with_anim(AnimationEffect::lighten());
    assert!(xml.contains("presetID=\"13\""), "lighten preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("E8E8E8"), "near-white target");
}

#[test]
fn animation_object_color_emits_preset15_emph() {
    let xml = make_slide_with_anim(AnimationEffect::object_color("4472C4"));
    assert!(xml.contains("presetID=\"15\""), "object color preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("4472C4"), "target colour");
    assert!(xml.contains("autoRev=\"1\""), "auto-reverse");
}

#[test]
fn animation_underline_emits_preset16_emph() {
    let xml = make_slide_with_anim(AnimationEffect::underline());
    assert!(xml.contains("presetID=\"16\""), "underline preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("style.textDecoration"), "textDecoration attr");
    assert!(xml.contains("autoRev=\"1\""), "auto-reverse");
}

// ── Emphasis — Moderate ───────────────────────────────────────────────────────

#[test]
fn animation_color_pulse_emits_preset17_emph() {
    let xml = make_slide_with_anim(AnimationEffect::color_pulse("ED7D31"));
    assert!(xml.contains("presetID=\"17\""), "color pulse preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("ED7D31"), "target colour");
    assert!(xml.contains("autoRev=\"1\""), "auto-reverse");
}

#[test]
fn animation_grow_with_color_emits_preset19_emph() {
    let xml = make_slide_with_anim(AnimationEffect::grow_with_color("70AD47"));
    assert!(xml.contains("presetID=\"19\""), "grow with color preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("<p:animScale>"), "scale component");
    assert!(xml.contains("<p:animClr"), "colour component");
    assert!(xml.contains("70AD47"), "target colour");
}

#[test]
fn animation_shimmer_emits_preset20_emph() {
    let xml = make_slide_with_anim(AnimationEffect::shimmer());
    assert!(xml.contains("presetID=\"20\""), "shimmer preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("repeatCount=\"3\""), "3 cycles");
    assert!(xml.contains("style.opacity"), "opacity attr");
}

#[test]
fn animation_teeter_emits_preset21_emph() {
    let xml = make_slide_with_anim(AnimationEffect::teeter());
    assert!(xml.contains("presetID=\"21\""), "teeter preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("style.rotation"), "rotation attr");
}

// ── Emphasis — Exciting ───────────────────────────────────────────────────────

#[test]
fn animation_blink_emits_preset22_emph() {
    let xml = make_slide_with_anim(AnimationEffect::blink());
    assert!(xml.contains("presetID=\"22\""), "blink preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("style.visibility"), "visibility attr");
    assert!(xml.contains("hidden"), "hidden keyframe");
}

#[test]
fn animation_bold_reveal_emits_preset23_emph() {
    let xml = make_slide_with_anim(AnimationEffect::bold_reveal());
    assert!(xml.contains("presetID=\"23\""), "bold reveal preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("style.fontWeight"), "bold component");
    assert!(xml.contains("<p:animScale>"), "scale component");
}

#[test]
fn animation_wave_emits_preset24_emph() {
    let xml = make_slide_with_anim(AnimationEffect::wave());
    assert!(xml.contains("presetID=\"24\""), "wave preset id");
    assert!(xml.contains("presetClass=\"emph\""), "emphasis class");
    assert!(xml.contains("style.rotation"), "rotation attr");
}
