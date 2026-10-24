# Changelog

All notable changes to deckmint, deckmint-math, and deckmint-wasm.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.9] - 2026-10-24

### Fixed
- `deckmint`: an empty text run (e.g. `TextRun::new("")` with `break_line()`, used to end a line, or a blank line) made PowerPoint offer to repair the file. `gen_xml_text_body_from_runs` wrote an `<a:endParaRPr>` in place of each empty run and another when closing the paragraph, but a paragraph may hold only one, as its last child. Empty runs now write nothing inline; a paragraph of only empty runs gives its closing `<a:endParaRPr>` the run's `sz`, so a blank line keeps its height. `scripts/check_pptx.py` gains `check_end_para_rpr`.
- `deckmint`: SVG pictures built via `add_image_svg`/`add_image_svg_from` (0.1.9) registered a structurally-valid `asvg:svgBlip` extension but PowerPoint declined to render it live — the picture stayed on its raster fallback until the user manually right-clicked → "Convert to Shape". Every genuine Office-authored SVG picture's `<a:blip>` carries **two** extensions in its `<a:extLst>` — `a14:useLocalDpi` (uri `{28A0092B-C50C-407E-A947-70E740481C1C}`) *then* `asvg:svgBlip` (uri `{96DAC541-7B7A-43D3-8B79-37D633B846F1}`) — confirmed against a real Excel-generated `xdr:blipFill`; deckmint (and, it turns out, upstream PptxGenJS's own `gen-xml.ts`) emitted only the second. "Convert to Shape" still worked because it reads the extension list directly regardless of what's alongside it, which is what made this look like a dangling-reference bug rather than a missing-sibling-extension one. `gen_xml_image_object` now emits `a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"` as the first `<a:ext>`, ahead of `svgBlip`, matching real Office output structure exactly. (The reported "wrong aspect ratio" was the raster fallback itself, not the frame it sits in — the picture's `<a:ext>` extent already agreed with the PNG fallback's own intrinsic pixel aspect; both were correct because `add_image_svg`'s auto-dimensioning reads the *raster's* pixel size, not the SVG's own `width`/`height`/`viewBox`. Extent and raster-intrinsic aspect are unrelated to this fix and unchanged.)
- `deckmint`: corrected the `rect_radius` docstring, which claimed the radius was a fraction of the shape's shorter side. The implementation has always treated `r` as an **absolute** radius in inches, converted to the OOXML `adj` guide value against the shape's *longer* side (`adj = round(r * EMU * 100_000 / max(cx, cy))`), so the same `r` reused across differently-proportioned shapes produces a different apparent corner radius. Behaviour is unchanged — only the docs were wrong — but a regression test now pins the true formula.
- `deckmint`: `TextOptionsBuilder::valign(...)` was silently dropped from the emitted `<a:bodyPr>` unless `.margin(...)` was *also* called on the same builder — `gen_xml_body_properties` only read `opts.valign` from inside the branch gated on `opts.margin.is_some()`. Any text or shape-with-text built with `.valign(AlignV::Middle)` alone got no `anchor` attribute at all, so PowerPoint defaulted it to top-anchored instead of centred. `valign` is now read independently of whether margins are set (an explicit `body_prop.anchor` override, when present, still takes precedence). Found in a downstream consumer whose pedagogical-badge labels sat visibly low in their pill shapes.
- `deckmint`: the `is_svg` branch of `gen_xml_image_object` emitted the `<asvg:svgBlip>` extension's raster-fallback `<a:blip r:embed>` as `rId{svg_rid - 1}` — assuming the caller had registered the fallback as the *immediately preceding* media relationship. No API path ever did that (`add_image`/`add_image_from` only ever register one media rel per call), so any SVG picture built through the public API referenced a fallback rId that was never defined, a dangling reference PowerPoint's repair dialog would have flagged on open. There was no working way to embed an SVG before this release.
- `deckmint`: `add_image`'s (and `SlideLayout`/`SlideMasterDef`'s) auto-dimension-from-bytes path filled in whichever of `w`/`h` was missing with the source image's *raw* intrinsic pixel size, ignoring the side that *was* given — `.w(4.0)` alone on a 2:1 image emitted the source's raw intrinsic height instead of scaling it to match the requested width, visibly stretching or squashing the picture. The missing side is now derived from the intrinsic aspect ratio scaled to the side that was given (a `Coord::Percent` side falls back to the old raw-intrinsic behaviour, since resolving a percentage to EMU needs the slide's own dimensions, unavailable at `add_image` time); giving both `w` and `h` still honours them exactly as-is.
- `deckmint-math`: `latex_to_omml` now repairs two `latex2mathml` 0.2.3 MathML serialisation bugs before parsing, so LaTeX containing a literal `<`/`&` (e.g. `x < 3`, `A \& B`) or a `\bigl(...\bigr)` sized-paren no longer fails with an "ill-formed document" XML error. `Node::Operator`/`Node::Letter`'s `Display` impls write unescaped characters into MathML text content, and `Node::SizedParen`'s `Display` impl closes `<mo …>` with a `</mro>` typo instead of `</mo>` — both are upstream bugs in the external `latex2mathml` crate (pinned to `"0.2"`), worked around internally in the new `xml_repair` module until they're fixed upstream. Found via a 10,348-span corpus sweep in a downstream consumer.

### Added
- `deckmint`: first-class SVG image embedding — `Slide::add_image_svg(svg_bytes, png_fallback_bytes, opts)` (+ `add_image_svg_base64`, `add_image_svg_from` builder variant with new `ImageOptionsBuilder::svg_fallback`/`svg_fallback_base64`), mirrored on `SlideLayout` and `SlideMasterDef` for chrome-tier SVGs. PowerPoint renders SVG via an `<asvg:svgBlip>` extension on a picture whose base `<a:blip r:embed>` must point at a real, independently-addressable raster image — `add_image_svg` registers *both* media parts as their own relationships (two rIds, two files under `ppt/media/`) rather than assuming adjacency, fixing the dangling-fallback bug described under Fixed. `ImageObject` gained `raster_fallback_rid: Option<u32>` (the SVG's own id stays in the existing `image_rid`); `remap_object_rids` (master/layout rId rebasing) and the SVG+PNG pair both participate in the existing package-wide media dedup pass, so an SVG reused across master/layout/slide only gets one physical copy of each part. Auto-dimensioning reads the **raster fallback's** pixel size (a plain pixel grid is a more reliable source of truth than parsing an arbitrary SVG's `viewBox`/`width`/`height`), through the same aspect-ratio-correct `resolve_image_size` helper the raster `add_image` fix above now shares. Verified end-to-end: a smoke deck with SVGs on a slide, a named layout, and the master opens cleanly in real PowerPoint (no repair dialog) and renders the SVG natively (crisp vector edges, not the blurry raster fallback) at every chrome tier.
- `deckmint`: `TextOptionsBuilder::underline(style)` sets a paragraph-level default underline style, mirroring the existing `.bold()`/`.italic()` pattern — a run only falls back to it when the run itself doesn't specify its own underline.
- `deckmint`: `Presentation::slide_layouts()` returns a read-only `&[SlideLayout]` slice, mirroring the existing `slide()`/`slide_count()` read-access pattern — for callers that build a whole named-layout set up front and then want to inspect what got defined.
- `deckmint`: `SlideLayout` is no longer a bare name/dimensions descriptor — it now carries the same shape-building surface as `SlideMasterDef` (`add_text`, `add_text_runs`, `add_shape`, `add_shape_with_text`, `add_image`, `add_image_base64`, `add_image_from`, `add_connector`, `add_group`, `add_table`), so a layout can carry its own chrome (badges, guide-rail outline boxes, an archetype icon) instead of always emitting an empty `<p:spTree>`. `Presentation::define_layout(...)` now returns `&mut SlideLayout` (matching `add_slide`'s "hand back the thing you just built" ergonomics, not `&mut Self` chaining) so callers can build straight onto it. New `Presentation::add_slide_with_layout(layout_name)` binds a slide to a specific named layout (panics if the name wasn't defined via `define_layout` first) — `ppt/slides/_rels/slideN.xml.rels`'s `rId1` now points at the right `slideLayoutM.xml` instead of every slide hardcoding layout 1. `SlideMasterDef` also gained `add_image_from` (was already on `Slide`, missing here) for API parity across all three chrome-bearing types. Layout media/hyperlink rIds are offset by 1 in `slideLayoutN.xml.rels` (rId1 is always reserved for that layout's slideMaster relationship) and layout media participates in the same package-wide dedup pass as master/slide media. Built for revision_atlas's deckgen, restructuring its generated decks to match a real house PowerPoint template's anatomy (chrome on the master + named layouts, slides carry only content) instead of re-stamping every chrome shape onto every slide.
- `deckmint-math`: `latex_to_omml_sized`/`mathml_to_pptx_omml_sized`/`postprocess::postprocess_omml_sized` — sized variants of the existing conversion entry points that stamp an explicit `sz="…"` onto every math run's `<a:rPr>` (and every structural block's `<m:ctrlPr>`), matching a caller-supplied point size. The existing unsized `latex_to_omml`/`mathml_to_pptx_omml`/`postprocess_omml` are unchanged (now thin wrappers passing `None`) and keep emitting sz-less `<a:rPr>` for back-compat. Fixes a font-parity bug in a downstream consumer: an OMML run's `<a:rPr>` previously carried a font face but no `sz` at all, so PowerPoint fell back to its own built-in default (18pt) instead of inheriting the surrounding paragraph's size the way an ordinary text run does — a 28pt prose line with inline `$…$` maths rendered its equation noticeably smaller than the text around it. `gen_xml_text_run` (the plain-run path) reads `run_opts.font_size.or(para_opts.font_size)` per run, but `<a:rPr>`-less OMML runs have no such fallback of their own; the fix is to make the OMML string itself carry the size, since deckmint's text-run emission short-circuits entirely for `equation_omml` runs (`s.push_str(omml); return s;`) and never reaches its own `sz=` logic for them.

### Changed
- `deckmint`: `scripts/check_pptx.py`'s `check_relationship_integrity` (orphaned rId / missing physical asset / duplicate rId cross-reference) now also runs over `ppt/slideLayouts/slideLayoutN.xml` and `ppt/slideMasters/slideMasterN.xml`, not just `ppt/slides/slideN.xml` — it previously would not have caught the SVG dangling-fallback bug above if the SVG had been placed on a layout or master instead of a slide.

## [0.1.7] - 2026-05-09

### Fixed
- Auto-firing animation cascades (where the first effect uses `.after_previous()` or `.with_previous()`) are now rewindable by PowerPoint's left arrow. Previously the leader's `.delay(N)` pushed `delay="N"` onto the step's outer `<p:cond>`, which PowerPoint treats as a non-navigable timeline event. The outer cond is now always `delay="indefinite"` (PowerPoint's "rewindable step" marker), and Auto steps add an `<p:cond evt="onBegin"><p:tn val="2"/></p:cond>` cond to auto-fire on slide enter; the user's leader delay flows into the leader's middle wrapper instead.
- A user-supplied `.delay(N)` on an `OnClick` leader was previously dropped (the outer cond was hardcoded to `indefinite`); it now lands on the leader's middle wrapper as intended.

### Changed
- `Step` no longer carries a separate `leader_delay_ms` field; the leader's delay is just `items[0].cumulative_delay_ms`.

## [0.1.6] - 2026-05-09

### Fixed
- Chained timed animations (`.after_previous()` / `.with_previous()`) now produce a single PowerPoint click step that left-arrow can rewind. Previously each chained effect emitted its own outer `<p:par>` with a unique `grpId`, which PowerPoint treated as separate non-rewindable steps.

### Changed
- `gen_xml_timing` grouping pass restructured: consecutive `WithPrevious` / `AfterPrevious` effects join the previous step instead of creating new ones. Effects in a step now share `grpId`, sit inside one outer `<p:par>` with per-effect cumulative-delay middle wrappers, and emit `<p:cond evt="onBegin"><p:tn val="2"/></p:cond>` on multi-effect steps to match PowerPoint's hand-authored structure.
- An animation marked `.after_previous()` / `.with_previous()` as the **first** animation on a slide now produces an "auto" step that fires on slide enter (rather than waiting for a click).

### Added
- Linter rule in `scripts/check_pptx.py` flagging `<p:seq prevAc="seek">` as schema-invalid (`ST_TLPreviousActionType` only allows `none` or `skipTimed`).
- 5 new integration tests covering chained-step `grpId` sharing, cumulative delay arithmetic, `withEffect` vs `afterEffect` node types, three-independent-clicks producing three `grpId`s, and leading `.after_previous()` producing an auto-step.

## [0.1.5] - 2026-05-09

### Fixed
- Reverted the `prevAc="seek"` attribute introduced in 0.1.4. `seek` is not a valid value for `prevAc` per OOXML (`ST_TLPreviousActionType` allows only `none` and `skipTimed`), and PowerPoint repaired any deck containing it.

## [0.1.4] - 2026-05-08 — Yanked

Buggy release: emitted `<p:seq prevAc="seek">` to try to enable left-arrow animation rewind. The attribute value is schema-invalid and caused PowerPoint to repair generated decks. Use 0.1.5 or later.

## [0.1.3] - 2026-04-09

### Fixed
- Out-of-range alpha values in shadow and glow XML output. Opacity values now normalise (>1.0 divided by 100, then clamped to 0.0–1.0) before scaling to the OOXML 0–100000 range.

### Added
- Linter rule for alpha values outside the 0–100000 range.

## [0.1.2] - 2026

### Added
- Per-crate `README.md` files for crates.io.

### Changed
- WASM demo run instructions; `pkg/` build output is now gitignored.

## [0.1.1] - 2026

### Added
- Ergonomic API improvements.
- Master slides support.
- 22 worked examples covering shapes, charts, tables, animations, transitions, equations, master slides, custom geometry, etc.
- Several repair-pattern fixes (animation XML, transitions, navigation hyperlinks, presentation sections).

## [0.1.0] - 2026

Initial release.

[0.1.9]: https://github.com/rebo/deckmint/releases/tag/v0.1.9
[0.1.7]: https://github.com/rebo/deckmint/releases/tag/v0.1.7
[0.1.6]: https://github.com/rebo/deckmint/releases/tag/v0.1.6
[0.1.5]: https://github.com/rebo/deckmint/releases/tag/v0.1.5
[0.1.4]: https://github.com/rebo/deckmint/releases/tag/v0.1.4
[0.1.3]: https://github.com/rebo/deckmint/releases/tag/v0.1.3
[0.1.2]: https://github.com/rebo/deckmint/releases/tag/v0.1.2
[0.1.1]: https://github.com/rebo/deckmint/releases/tag/v0.1.1
[0.1.0]: https://github.com/rebo/deckmint/releases/tag/v0.1.0
