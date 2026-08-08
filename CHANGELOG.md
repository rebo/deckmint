# Changelog

All notable changes to deckmint, deckmint-math, and deckmint-wasm.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed
- `deckmint-math`: `latex_to_omml` now repairs two `latex2mathml` 0.2.3 MathML serialisation bugs before parsing, so LaTeX containing a literal `<`/`&` (e.g. `x < 3`, `A \& B`) or a `\bigl(...\bigr)` sized-paren no longer fails with an "ill-formed document" XML error. `Node::Operator`/`Node::Letter`'s `Display` impls write unescaped characters into MathML text content, and `Node::SizedParen`'s `Display` impl closes `<mo …>` with a `</mro>` typo instead of `</mo>` — both are upstream bugs in the external `latex2mathml` crate (pinned to `"0.2"`), worked around internally in the new `xml_repair` module until they're fixed upstream. Found via a 10,348-span corpus sweep in a downstream consumer.
- `deckmint`: corrected the `rect_radius` docstring, which claimed the radius was a fraction of the shape's shorter side. The implementation has always treated `r` as an **absolute** radius in inches, converted to the OOXML `adj` guide value against the shape's *longer* side (`adj = round(r * EMU * 100_000 / max(cx, cy))`), so the same `r` reused across differently-proportioned shapes produces a different apparent corner radius. Behaviour is unchanged — only the docs were wrong — but a regression test now pins the true formula.

### Added
- `deckmint`: `TextOptionsBuilder::underline(style)` sets a paragraph-level default underline style, mirroring the existing `.bold()`/`.italic()` pattern — a run only falls back to it when the run itself doesn't specify its own underline.

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

[0.1.7]: https://github.com/rebo/deckmint/releases/tag/v0.1.7
[0.1.6]: https://github.com/rebo/deckmint/releases/tag/v0.1.6
[0.1.5]: https://github.com/rebo/deckmint/releases/tag/v0.1.5
[0.1.4]: https://github.com/rebo/deckmint/releases/tag/v0.1.4
[0.1.3]: https://github.com/rebo/deckmint/releases/tag/v0.1.3
[0.1.2]: https://github.com/rebo/deckmint/releases/tag/v0.1.2
[0.1.1]: https://github.com/rebo/deckmint/releases/tag/v0.1.1
[0.1.0]: https://github.com/rebo/deckmint/releases/tag/v0.1.0
