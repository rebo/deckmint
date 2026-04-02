# deckmint — Developer Notes for Claude

## Repair Issue Protocol

**Whenever a PowerPoint repair issue is found, you MUST do two things:**

1. Fix the XML generator (`deckmint/src/xml/slide_xml.rs` or relevant source)
2. Add a linter check to `scripts/check_pptx.py` so the pattern is caught automatically in future

This applies whether the repair is discovered by opening a generated `.pptx` in PowerPoint, by running the linter, or by examining a repaired file.

---

## Repair Investigation Workflow

When PowerPoint repairs a file it saves a repaired copy. Compare it against a fresh build:

```bash
# Unzip both files
mkdir /tmp/orig /tmp/repaired
cd /tmp/orig   && unzip /path/to/original.pptx
cd /tmp/repaired && unzip /path/to/repaired.pptx

# Format and diff a slide
xmllint --format /tmp/orig/ppt/slides/slide1.xml     > /tmp/orig.xml
xmllint --format /tmp/repaired/ppt/slides/slide1.xml > /tmp/repaired.xml
diff /tmp/orig.xml /tmp/repaired.xml
```

Focus on **structural** differences (element names, attribute names, element nesting). Ignore:
- Attribute ordering (XML parsers may reorder)
- ID number differences (PowerPoint resequences all IDs)
- Whitespace-only diffs

---

## Adding a Check to the Linter

File: `scripts/check_pptx.py`

The linter walks all slide XML files and calls check functions. To add a new check:

1. Add a function with this signature:
   ```python
   def check_my_pattern(slide_path: str, tree) -> list[str]:
       """One-line description of what this catches."""
       issues = []
       # ... use tree.findall / xpath ...
       return issues
   ```

2. Wire it into `main()` in the existing check loop:
   ```python
   issues += check_my_pattern(path, tree)
   ```

The function receives the file path (for error messages) and the parsed `lxml` ElementTree root. Return a list of human-readable issue strings (empty list = no issues).

---

## Known Repair Patterns (History)

These have already been fixed and are checked by the linter:

| Pattern | Fix | Check function |
|---|---|---|
| `fill="hold"` on `<p:cBhvr>` | Move to `<p:cTn>` (only for `animClr`/`set`; omit for plain `anim`) | `check_animation_xml` |
| `<p:clrVal>` wrapping `<a:srgbClr>` inside `<p:to>` in `<p:animClr>` | Remove `<p:clrVal>`; bare `<a:srgbClr>` goes directly in `<p:to>` | `check_animation_xml` |
| `<a:hslClr>` in `<p:by>` of `<p:animClr>` | Replace with `<p:hsl h="..." s="..." l="..."/>` (p: namespace) | `check_animation_xml` |
| Missing `dir="cw"` on `<p:animClr>` | Add `dir="cw"` attribute | `check_animation_xml` |
| `animBg` missing from `<p:bldP>` | Always emit `animBg="1"` on every `<p:bldP>` | `check_animation_xml` |
| `stIdx`/`endIdx` on `<p:charRg>`/`<p:pRg>` | Use `st` and `end` (correct `CT_IndexRange` attribute names) | `check_animation_xml` |
| `<c:tickMark>` in chart axis XML | Use `<c:majorTickMark>` + `<c:minorTickMark>` — `c:tickMark` is not a valid OOXML element | `check_chart_xml` |
| `<a:noFill/>` before `<a:ln>` inside a line/scatter series `<c:spPr>` | Remove `<a:noFill/>` — the series shape has no fill area; PowerPoint removes it during repair | `check_chart_xml` |
| Missing chart-level `<c:dLbls>` inside chart type element | Add `<c:dLbls><c:showLegendKey val="0"/>…</c:dLbls>` after series, before `<c:axId>`/`<c:firstSliceAng>` in every chart type | `check_chart_xml` |
| Wrong axis positions in horizontal bar chart (`barDir="bar"`) | catAx must use `axPos="l"`, valAx must use `axPos="b"` — column chart positions are reversed | `check_chart_xml` |
| `<c:majorGridlines>` in catAx | Move to valAx only — horizontal category-axis gridlines are not standard and PowerPoint removes them | `check_chart_xml` |
| Missing `<c:marker>`/`<c:smooth>` chart-level elements in lineChart | Add `<c:marker val="0/1"/>` and `<c:smooth val="0/1"/>` after chart-level `<c:dLbls>`, before `<c:axId>` | (implicitly checked via dLbls position) |
| Series `<c:dLbls>` emitted after `<c:cat>`/`<c:val>` | Move dLbls to before cat/val — CT_BarSer/CT_LineSer content model requires this order | `check_chart_xml` |
| Chart-level `<c:smooth val="1"/>` in lineChart | Always emit `val="0"` at chart level; per-series `<c:smooth>` handles actual smoothing | (avoid setting to 1) |
| `<p:bgPr>` with solidFill/gradFill color background missing `<a:effectLst/>` | Add `<a:effectLst/>` after any fill inside `<p:bgPr>` — required for both image and color backgrounds | `check_bgpr_effectlst` |
| `xmlns:p14` declared on root `<p:presentation>` element | Move namespace inline to `<p14:sectionLst xmlns:p14="...">` — root-level declaration causes PowerPoint to strip the section list on repair | `check_presentation_sections` |
| `<p14:sldId r:id="rIdN"/>` in section list | Use `<p14:sldId id="NUM"/>` with the numeric slide ID (e.g. 256), not the relationship ID | `check_presentation_sections` |
| Missing `<p15:sldGuideLst/>` companion after `<p14:sectionLst>` | Add `<p:ext uri="{EFAFB233-...}"><p15:sldGuideLst xmlns:p15="..."/></p:ext>` after the section ext block | `check_presentation_sections` |
| `<p:transition advClick="1">` | Omit `advClick` when true — it is the OOXML schema default | `check_transition_xml` |
| `<p:push dir="l"/>` (and other directional transitions) | Omit `dir` attribute when it equals `"l"` (left is the OOXML default direction) | `check_transition_xml` |
| `<a:prstGeom prst="elbow">` / `prst="curve"` in connectors | Use `"bentConnector3"` / `"curvedConnector3"` (correct OOXML preset names) | `check_transition_xml` |
| `<a:hlinkClick action="ppaction://hlinkshowjump?...">` missing `r:id` | Add `r:id=""` — CT_Hyperlink requires the attribute even with no relationship | `check_nav_hyperlink_rid` |

---

## Running the Linter

```bash
# After building an example
cargo run --example features
python3 scripts/check_pptx.py features.pptx
```

Expected output on a clean file: `No repair-pattern issues found.`
