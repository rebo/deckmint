# deckmint Feature Comparison

Comparison of the TypeScript PptxGenJS library versus the Rust `deckmint` crate.

| Feature | TypeScript | Rust | Notes |
|---------|-----------|------|-------|
| **Core** | | | |
| Create presentation | ✅ | ✅ | |
| Set title/author/subject/company | ✅ | ✅ | |
| Custom layout dimensions | ✅ | ✅ | `PresLayout` |
| Multiple slides | ✅ | ✅ | |
| Write to bytes / file | ✅ | ✅ | `write()` / `write_to_file()` |
| WASM support | ✅ | ✅ | `deckmint-wasm` crate |
| **Text** | | | |
| Basic text box | ✅ | ✅ | |
| Bold / italic / strike | ✅ | ✅ | |
| Font size / face / color | ✅ | ✅ | |
| Underline (single) | ✅ | ✅ | |
| Underline styles (15 variants) | ✅ | ✅ | `dbl`, `dash`, `dotted`, `wavy`, etc. |
| Text alignment (H + V) | ✅ | ✅ | |
| RTL text | ✅ | ✅ | |
| Line spacing / paragraph spacing | ✅ | ✅ | |
| Bullets (default/numbered/char) | ✅ | ✅ | |
| Superscript / subscript | ✅ | ✅ | |
| Char spacing (kerning) | ✅ | ✅ | |
| Highlight color | ✅ | ✅ | `<a:highlight>` |
| Text glow effect | ✅ | ✅ | `GlowProps` |
| Text outline (stroke) | ✅ | ✅ | `TextOutlineProps` |
| Text direction (vert/horz) | ✅ | ✅ | `vert` attribute on `<a:bodyPr>` |
| Soft line break (`<a:br>`) | ✅ | ✅ | `soft_break_before` on `TextRun` |
| Tab stops | ✅ | ✅ | `TabStop` + `tab_stops` on `TextOptions` |
| Rich text runs | ✅ | ✅ | `add_text_runs()` |
| Text fill (background) | ✅ | ✅ | |
| Text wrap | ✅ | ✅ | |
| Text fit (shrink/resize) | ✅ | ✅ | |
| Text hyperlinks | ✅ | ✅ | URL and slide-jump |
| Indent level | ✅ | ✅ | |
| Speaker notes | ✅ | ✅ | |
| **Shapes** | | | |
| Basic shapes (150+ types) | ✅ | ✅ | `ShapeType` enum |
| Shape fill color | ✅ | ✅ | |
| Shape line (color/width/dash) | ✅ | ✅ | |
| Shape rotation / flip | ✅ | ✅ | |
| Rounded rectangle radius | ✅ | ✅ | |
| Shape shadow | ✅ | ✅ | outer/inner shadow |
| Shadow rotateWithShape | ✅ | ✅ | `rotate_with_shape` field |
| Shape with text | ✅ | ✅ | `add_shape_with_text()` |
| Shape hyperlink | ✅ | ✅ | `hyperlink` field on `ShapeOptions` |
| Arrow heads | ✅ | ✅ | `begin_arrow_type`/`end_arrow_type` |
| **Images** | | | |
| Embed PNG/JPEG/GIF/SVG | ✅ | ✅ | |
| Image sizing (cover/contain/crop) | ✅ | ✅ | |
| Image transparency | ✅ | ✅ | |
| Rounded (ellipse) image | ✅ | ✅ | |
| Image auto-dimension detection | ✅ | ✅ | `imagesize` crate (96 DPI) |
| Image hyperlink | ✅ | ✅ | |
| Image alt text | ✅ | ✅ | |
| Image from base64 | ✅ | ✅ | `add_image_base64()` |
| **Tables** | | | |
| Basic table | ✅ | ✅ | |
| Column widths | ✅ | ✅ | `col_w` |
| Per-row heights | ✅ | ✅ | `row_h` (per-row wired) |
| Cell fill / color | ✅ | ✅ | |
| Cell font (size/face/bold/italic) | ✅ | ✅ | |
| Cell alignment (H + V) | ✅ | ✅ | |
| Cell margins | ✅ | ✅ | cell-level and table-level default |
| Cell borders | ✅ | ✅ | per-side with color |
| Cell colspan / rowspan | ✅ | ✅ | |
| Cell hyperlink | ✅ | ✅ | `hyperlink` on `TableCellProps` |
| Table shadow | ✅ | ✅ | |
| Table auto-pagination | ✅ | ❌ | Complex feature, Phase 9 |
| **Slide** | | | |
| Background color | ✅ | ✅ | |
| Background image | ✅ | ✅ | `set_background_image()` |
| Hide slide | ✅ | ✅ | |
| Slide hyperlinks (URL + jump) | ✅ | ✅ | |
| **Slide Master** | | | |
| Define slide master background | ✅ | ✅ | `SlideMasterDef.background_color` |
| Define slide master objects | ✅ | ✅ | `SlideMasterDef.objects` |
| Define custom slide layout | ✅ | ✅ | `define_layout()` with dimensions |
| Multiple slide masters | ✅ | ❌ | Single master only |
| **Packaging** | | | |
| Valid OOXML ZIP structure | ✅ | ✅ | |
| Content types | ✅ | ✅ | |
| Relationships | ✅ | ✅ | |
| Theme XML | ✅ | ✅ | custom font faces |
| Notes master | ✅ | ✅ | |

## Legend
- ✅ Fully implemented
- ❌ Not implemented
