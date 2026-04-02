use crate::error::PptxError;
use crate::objects::SlideObject;
use crate::packaging::assemble_pptx;
use crate::slide::Slide;
use crate::types::PresLayout;

/// A slide layout descriptor
#[derive(Debug, Clone)]
pub struct SlideLayout {
    pub name: String,
    /// Optional custom width in inches (defaults to presentation width)
    pub width: Option<f64>,
    /// Optional custom height in inches (defaults to presentation height)
    pub height: Option<f64>,
}

/// Definition of the slide master
#[derive(Debug, Clone, Default)]
pub struct SlideMasterDef {
    pub title: String,
    /// Background fill color as hex (no #)
    pub background_color: Option<String>,
    /// Objects to render on the master slide
    pub objects: Vec<SlideObject>,
}

/// Optional presentation theme (font faces)
#[derive(Debug, Clone, Default)]
pub struct ThemeProps {
    pub head_font_face: Option<String>,
    pub body_font_face: Option<String>,
}

/// A named section grouping slides
#[derive(Debug, Clone)]
pub struct SectionDef {
    /// Visible section name
    pub name: String,
    /// 1-based slide number of the first slide in this section
    pub start_slide: u32,
}

/// Top-level presentation builder
#[derive(Debug)]
pub struct Presentation {
    pub author: String,
    pub company: String,
    pub title: String,
    pub subject: String,
    pub revision: String,
    pub rtl_mode: bool,
    pub layout: PresLayout,
    pub theme: Option<ThemeProps>,
    pub(crate) slides: Vec<Slide>,
    pub(crate) slide_layouts: Vec<SlideLayout>,
    pub(crate) master: Option<SlideMasterDef>,
    /// Named sections (groupings visible in the slide panel)
    pub(crate) sections: Vec<SectionDef>,
    slide_id_counter: u32,
}

impl Default for Presentation {
    fn default() -> Self {
        Presentation {
            author: String::new(),
            company: String::new(),
            title: String::new(),
            subject: String::new(),
            revision: "1".to_string(),
            rtl_mode: false,
            layout: PresLayout::default(),
            theme: None,
            slides: Vec::new(),
            slide_layouts: vec![SlideLayout { name: "DEFAULT".to_string(), width: None, height: None }],
            master: None,
            sections: Vec::new(),
            slide_id_counter: 0,
        }
    }
}

impl Presentation {
    /// Create a new blank presentation with 16:9 layout
    pub fn new() -> Self {
        Presentation::default()
    }

    /// Create a presentation with a custom layout
    pub fn with_layout(layout: PresLayout) -> Self {
        Presentation { layout, ..Default::default() }
    }

    /// Set the author metadata
    pub fn author(mut self, a: impl Into<String>) -> Self { self.author = a.into(); self }
    /// Set the company metadata
    pub fn company(mut self, c: impl Into<String>) -> Self { self.company = c.into(); self }
    /// Set the title metadata
    pub fn title(mut self, t: impl Into<String>) -> Self { self.title = t.into(); self }
    /// Set the subject metadata
    pub fn subject(mut self, s: impl Into<String>) -> Self { self.subject = s.into(); self }
    /// Set the revision metadata
    pub fn revision(mut self, r: impl Into<String>) -> Self { self.revision = r.into(); self }
    /// Enable right-to-left mode
    pub fn rtl(mut self) -> Self { self.rtl_mode = true; self }
    /// Set a custom theme
    pub fn theme(mut self, t: ThemeProps) -> Self { self.theme = Some(t); self }

    // ─── Layout / Master management ─────────────────────────

    /// Define a custom slide layout with optional dimensions in inches.
    /// The layout will be added to the layout list used by slides.
    pub fn define_layout(&mut self, name: impl Into<String>, width: Option<f64>, height: Option<f64>) -> &mut Self {
        self.slide_layouts.push(SlideLayout { name: name.into(), width, height });
        self
    }

    /// Set the slide master definition. Objects and background defined here
    /// are rendered into `ppt/slideMasters/slideMaster1.xml`.
    pub fn define_master(&mut self, def: SlideMasterDef) -> &mut Self {
        self.master = Some(def);
        self
    }

    // ─── Slide management ───────────────────────────────────

    /// Add a named section starting at the next slide to be added.
    /// Sections are visible in the slide panel of PowerPoint / LibreOffice.
    pub fn add_section(&mut self, name: impl Into<String>) -> &mut Self {
        let next_slide_num = self.slide_id_counter + 1;
        self.sections.push(SectionDef { name: name.into(), start_slide: next_slide_num });
        self
    }

    /// Add a new blank slide and return a mutable reference to it.
    pub fn add_slide(&mut self) -> &mut Slide {
        self.slide_id_counter += 1;
        let slide_num = self.slide_id_counter;
        let slide_id = 255 + slide_num; // OOXML spec: first slide ID is 256
        let r_id = slide_num + 1;       // rId1 = slide master, rId2 = slide 1, etc.
        let slide = Slide::new(slide_num, slide_id, r_id);
        self.slides.push(slide);
        self.slides.last_mut().unwrap()
    }

    /// Return the number of slides
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Get a reference to a slide by 0-based index
    pub fn slide(&self, idx: usize) -> Option<&Slide> {
        self.slides.get(idx)
    }

    /// Get a mutable reference to a slide by 0-based index
    pub fn slide_mut(&mut self, idx: usize) -> Option<&mut Slide> {
        self.slides.get_mut(idx)
    }

    // ─── Export ─────────────────────────────────────────────

    /// Assemble the presentation into a ZIP buffer (`Vec<u8>`).
    /// This is synchronous and works on both native and WASM.
    pub fn write(&self) -> Result<Vec<u8>, PptxError> {
        assemble_pptx(self)
    }

    /// Write the presentation to a file (native only).
    #[cfg(not(target_arch = "wasm32"))]
    pub fn write_to_file(&self, path: &str) -> Result<(), PptxError> {
        let bytes = self.write()?;
        std::fs::write(path, &bytes).map_err(PptxError::Io)
    }

    /// Assemble and validate the presentation.
    ///
    /// Returns the PPTX bytes if validation passes, or a
    /// [`PptxError::ValidationFailed`] listing all structural issues.
    pub fn write_validated(&self) -> Result<Vec<u8>, PptxError> {
        let bytes = self.write()?;
        let issues = crate::validate::validate(&bytes);
        let errors: Vec<_> = issues.into_iter()
            .filter(|i| i.severity == crate::validate::Severity::Error)
            .collect();
        if !errors.is_empty() {
            return Err(PptxError::ValidationFailed(errors));
        }
        Ok(bytes)
    }
}
