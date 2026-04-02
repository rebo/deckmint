#!/usr/bin/env python3
"""
check_pptx.py — structural repair-pattern validator for deckmint output.

Catches issues that pass OOXML schema validation but still trigger PowerPoint's
"needs repair" dialog. Run after @xarsh/ooxml-validator for full coverage.

Usage:
    python3 scripts/check_pptx.py <file.pptx>

Exit code 0 = clean, 1 = issues found.
"""

import sys
import zipfile
import re
from xml.etree import ElementTree as ET

NS = {
    "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p":  "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel":"http://schemas.openxmlformats.org/package/2006/relationships",
}

SLIDE_LAYOUT_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
NOTES_SLIDE_TYPE  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
SLIDE_MASTER_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
SLIDE_TYPE        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
IMAGE_TYPE        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
HYPERLINK_TYPE    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
OFFICE_DOC_TYPE   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"

errors = []

def err(msg):
    errors.append(msg)
    print(f"  FAIL  {msg}")

def ok(msg):
    print(f"  ok    {msg}")


# ─── Tier 0: Repair-pattern checks (PowerPoint semantic conventions) ──────────

def check_slide_rels(zf, slide_name):
    """
    rId1 must be slideLayout, rId2 must be notesSlide.
    User content (hyperlinks, images) must start at rId3+.
    """
    rels_path = slide_name.replace("slides/", "slides/_rels/") + ".rels"
    if rels_path not in zf.namelist():
        err(f"{rels_path}: missing rels file")
        return

    xml = zf.read(rels_path).decode("utf-8")
    root = ET.fromstring(xml)

    rid_to_type = {}
    for rel in root:
        rid = rel.get("Id", "")
        typ = rel.get("Type", "")
        rid_to_type[rid] = typ

    layout_rid = next((r for r, t in rid_to_type.items() if t == SLIDE_LAYOUT_TYPE), None)
    notes_rid  = next((r for r, t in rid_to_type.items() if t == NOTES_SLIDE_TYPE), None)

    if layout_rid != "rId1":
        err(f"{rels_path}: slideLayout must be rId1 (found {layout_rid})")
    else:
        ok(f"{rels_path}: slideLayout=rId1")

    if notes_rid != "rId2":
        err(f"{rels_path}: notesSlide must be rId2 (found {notes_rid})")
    else:
        ok(f"{rels_path}: notesSlide=rId2")

    # User content must not use rId1 or rId2
    for rid, typ in rid_to_type.items():
        if rid in ("rId1", "rId2") and typ not in (SLIDE_LAYOUT_TYPE, NOTES_SLIDE_TYPE):
            err(f"{rels_path}: {rid} used by user content ({typ.split('/')[-1]}) — must be rId3+")


def check_notes_slide(zf, notes_path):
    """
    Empty notes slides must not contain an empty <a:r> text run.
    An empty paragraph should just have <a:endParaRPr>.
    """
    xml = zf.read(notes_path).decode("utf-8")

    # Check for <a:t></a:t> (empty text element in a run)
    if re.search(r'<a:t\s*/>', xml) or "<a:t></a:t>" in xml:
        # Only an error if there's no actual content
        content = re.findall(r'<a:t[^>]*>(.*?)</a:t>', xml)
        empty_runs = [c for c in content if c.strip() == ""]
        if empty_runs:
            err(f"{notes_path}: contains {len(empty_runs)} empty <a:t> run(s) — omit <a:r> when notes text is empty")
        else:
            ok(f"{notes_path}: no empty text runs")
    else:
        ok(f"{notes_path}: no empty text runs")


def check_empty_elements(zf, xml_path):
    """
    Empty <a:bodyPr> and <a:rPr> elements should be self-closing.
    Open/close tags with no children trigger PowerPoint repair.
    """
    xml = zf.read(xml_path).decode("utf-8")

    # Detect open+close empty tags for known offenders
    patterns = [
        (r'<(a:bodyPr[^>]*)>\s*</a:bodyPr>', "a:bodyPr"),
        (r'<(a:rPr[^>]*)>\s*</a:rPr>',       "a:rPr"),
    ]
    for pattern, tag in patterns:
        matches = re.findall(pattern, xml)
        if matches:
            err(f"{xml_path}: {len(matches)} open/close empty <{tag}> — should be self-closing")
        else:
            ok(f"{xml_path}: <{tag}> self-closing ok")


def check_shadow_defaults(zf, xml_path):
    """
    <a:outerShdw> should not include default attributes sx/sy/kx/ky
    or any rotWithShape value (both "0" and "1" are stripped by PowerPoint on repair).
    """
    xml = zf.read(xml_path).decode("utf-8")
    if "outerShdw" not in xml:
        return

    bad_attrs = ["sx=", "sy=", "kx=", "ky=", "rotWithShape="]
    for m in re.finditer(r'<a:outerShdw[^>]*>', xml):
        tag = m.group(0)
        found = [a for a in bad_attrs if a in tag]
        if found:
            err(f"{xml_path}: <a:outerShdw> contains attrs that should be omitted: {found}")
            return
    ok(f"{xml_path}: <a:outerShdw> attributes ok")


def check_animation_xml(zf, xml_path):
    """
    Animation (p:timing) repair-pattern checks — accumulated from PowerPoint repair sessions.

    Checks (each maps to a known repair trigger):

    1. charRg/pRg must use st= and end= (not stIdx=/endIdx=)
       Source: rep6 — PowerPoint replaced stIdx/endIdx with sentinel 0xFFFFFFFF.

    2. <p:cBhvr> must NOT have a fill= attribute.
       CT_TLCommonBehaviorData has no fill attr; fill belongs on <p:cTn>.
       Source: rep5 — PowerPoint stripped fill= from cBhvr.

    3. <p:animClr> must carry dir= (typically "cw").
       Source: rep5 — PowerPoint added dir="cw" on repair.

    4. <p:to> / <p:by> inside <p:animClr> must NOT wrap colors in <p:clrVal>.
       The type is CT_TLByAnimateColorTransform, not CT_TLAnimVariant.
       Source: rep5 — PowerPoint removed <p:clrVal> wrapper.

    5. <p:by> inside <p:animClr clrSpc="hsl"> must use <p:hsl> or <p:rgb>,
       NOT <a:hslClr> (DrawingML absolute color — wrong namespace/element).
       Source: rep5 — PowerPoint replaced <a:hslClr> with <p:rgb -100000>.

    6. <p:bldP> animBg rules:
       - Whole-shape/whole-paragraph animations: must have animBg="1"
       - Sub-target (charRg/pRg) animations: must NOT have animBg="1"; use uiExpand="1"
       animBg="1" on a charRg bldP causes PowerPoint to apply the colour to the whole
       paragraph instead of just the targeted characters.
       Source: rep5 (animBg required for whole-shape); user testing (animBg wrong for charRg).
    """
    xml = zf.read(xml_path).decode("utf-8")
    if "<p:timing>" not in xml:
        return  # no animation on this slide

    issues = []

    # 1. charRg / pRg must use st= / end=, not stIdx= / endIdx=
    if re.search(r'<p:(?:charRg|pRg)\b[^>]*(?:stIdx|endIdx)=', xml):
        issues.append("charRg/pRg uses stIdx=/endIdx= — must be st=/end= (CT_IndexRange attributes)")

    # 2. <p:cBhvr> must not carry fill=
    if re.search(r'<p:cBhvr\b[^>]*\bfill=', xml):
        issues.append("<p:cBhvr fill=…> — fill attribute is invalid on CT_TLCommonBehaviorData; move it to <p:cTn>")

    # 3. <p:animClr> must have dir= attribute
    for m in re.finditer(r'<p:animClr\b([^>]*)>', xml):
        attrs = m.group(1)
        if 'dir=' not in attrs:
            issues.append("<p:animClr> missing dir= attribute (expected dir=\"cw\" or dir=\"ccw\")")
            break  # one report is enough

    # 4. <p:clrVal> inside <p:to> or <p:by> within an animClr block
    # Detect the pattern <p:to><p:clrVal> or <p:by><p:clrVal>
    if re.search(r'<p:(?:to|by|from)>\s*<p:clrVal>', xml):
        issues.append("<p:to/by/from><p:clrVal>…</p:clrVal> — color inside animClr from/to/by must be a bare "
                      "color element (<a:srgbClr> etc.), not wrapped in <p:clrVal> (CT_TLByAnimateColorTransform "
                      "is not CT_TLAnimVariant)")

    # 5. <a:hslClr> inside <p:by> (wrong element; should be <p:hsl>)
    if re.search(r'<p:by>\s*(?:<p:clrVal>\s*)?<a:hslClr\b', xml):
        issues.append("<p:by><a:hslClr> — use <p:hsl h=… s=… l=…/> (p: namespace) for HSL adjustments "
                      "in animClr; a:hslClr is an absolute DrawingML color, not a relative delta")

    # 6. <p:bldP> animBg rules (uses the parsed tree for correct nesting):
    #    - Whole-shape / whole-paragraph animations: must have animBg="1"
    #    - Sub-target animations (charRg / pRg): must NOT have animBg="1"; use uiExpand="1"
    #    animBg="1" on a charRg bldP causes PowerPoint to colour the whole paragraph.
    try:
        tree = ET.fromstring(xml.encode("utf-8"))
        p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"

        # Walk all <p:cTn> that carry a grpId (these are the click-effect nodes).
        # For each, check whether any descendant is a <p:charRg> or <p:pRg>.
        sub_target_grp_ids_tree = set()
        for ctn in tree.iter(f"{{{p_ns}}}cTn"):
            grp_id = ctn.get("grpId")
            if grp_id is None:
                continue
            # Check all descendants of this cTn for txEl with charRg/pRg
            for child in ctn.iter():
                if child.tag in (f"{{{p_ns}}}charRg", f"{{{p_ns}}}pRg"):
                    sub_target_grp_ids_tree.add(grp_id)
                    break

        for bld in tree.iter(f"{{{p_ns}}}bldP"):
            grp_id = bld.get("grpId", "?")
            has_anim_bg = bld.get("animBg") == "1"
            is_sub = grp_id in sub_target_grp_ids_tree

            if is_sub and has_anim_bg:
                issues.append(
                    f"<p:bldP grpId=\"{grp_id}\" animBg=\"1\"> — sub-target (charRg/pRg) animations "
                    "must NOT have animBg=\"1\"; use uiExpand=\"1\" (animBg causes whole-paragraph colouring)")
                break
            if not is_sub and not has_anim_bg:
                issues.append(
                    f"<p:bldP grpId=\"{grp_id}\"> missing animBg=\"1\" — "
                    "whole-shape animations require animBg=\"1\"")
                break
    except ET.ParseError:
        pass  # malformed XML will be caught by other checks

    if issues:
        for issue in issues:
            err(f"{xml_path}: {issue}")
    else:
        ok(f"{xml_path}: animation XML ok")


def check_presentation_sections(zf):
    """
    Checks for p14:sectionLst correctness:
    1. xmlns:p14 must be inline on <p14:sectionLst>, not on root <p:presentation>
    2. <p14:sldId> must use id="NUM" (numeric slide ID), NOT r:id="rIdN"
    3. A companion <p15:sldGuideLst/> ext block must follow the section ext block
    """
    if "ppt/presentation.xml" not in zf.namelist():
        return
    xml = zf.read("ppt/presentation.xml").decode("utf-8")
    if "sectionLst" not in xml:
        ok("ppt/presentation.xml: no sections (ok)")
        return
    issues = []
    # Check 1: p14 namespace must NOT appear on root <p:presentation>
    root_line = xml.split(">", 1)[0]
    if 'xmlns:p14' in root_line:
        issues.append("xmlns:p14 declared on root <p:presentation> — must be inline on <p14:sectionLst>")
    elif 'xmlns:p14' not in xml:
        issues.append("<p14:sectionLst> present but xmlns:p14 never declared")
    # Check 2: <p14:sldId> must use id= not r:id=
    if 'p14:sldId r:id=' in xml:
        issues.append("<p14:sldId> uses r:id= attribute — must use id= (numeric slide ID)")
    # Check 3: must have companion p15:sldGuideLst
    if 'sldGuideLst' not in xml:
        issues.append("missing <p15:sldGuideLst/> companion extension after <p14:sectionLst>")
    if issues:
        for issue in issues:
            err(f"ppt/presentation.xml: {issue}")
    else:
        ok("ppt/presentation.xml: p14:sectionLst format ok")


def check_transition_xml(zf, xml_path):
    """
    Slide transition repair-pattern checks.

    1. <p:transition advClick="1"> — advClick="1" is the OOXML default; emitting it
       triggers repair on some PowerPoint versions.
       Source: rep-x1 — PowerPoint stripped advClick="1" from all transitions.

    2. Directional transition elements (push, wipe, cover, etc.) must NOT emit
       dir="l" since "l" (left) is the OOXML default direction.
       Source: rep-x1 — PowerPoint removed dir="l" from <p:push dir="l"/>.

    3. Connector shapes must use the correct OOXML preset geometry names:
       - Elbow connectors:  bentConnector3  (not "elbow")
       - Curved connectors: curvedConnector3 (not "curve")
       Incorrect prst values cause the geometry to fail to parse.
       Source: rep-x1 — PowerPoint corrected "elbow"/"curve" → "line" (fallback).
    """
    xml = zf.read(xml_path).decode("utf-8")

    issues = []

    # 1. advClick="1" should not be emitted (it's the default)
    if re.search(r'<p:transition\b[^>]*advClick="1"', xml):
        issues.append('<p:transition advClick="1"> — advClick="1" is the OOXML default; omit it')

    # 2. dir="l" on directional transition elements should not be emitted
    dir_elements = ["push", "wipe", "cover", "uncover", "pan", "ferris", "gallery",
                    "conveyor", "doors", "box", "strips"]
    for el in dir_elements:
        if re.search(rf'<p:{el}\b[^>]*dir="l"', xml):
            issues.append(f'<p:{el} dir="l"> — "l" is the default direction; omit dir attribute')
            break  # one report is enough

    # 3. <p:checkerboard> is not a valid OOXML element — use <p:checker>
    if "<p:checkerboard" in xml:
        issues.append('<p:checkerboard> — invalid element; use <p:checker> for checkerboard transition')

    # 4. <p:flash> is not in the base OOXML schema — use <p14:flash> in mc:AlternateContent
    if re.search(r'<p:flash\s*/?\s*>', xml) and 'p14:flash' not in xml:
        issues.append('<p:flash> — not a valid p: element; use <p14:flash> in mc:AlternateContent')

    # 5. Invalid connector preset geometry names
    for bad_prst in ("elbow", "curve"):
        if re.search(rf'<a:prstGeom\b[^>]*prst="{bad_prst}"', xml):
            correct = "bentConnector3" if bad_prst == "elbow" else "curvedConnector3"
            issues.append(
                f'<a:prstGeom prst="{bad_prst}"> — invalid connector preset; '
                f'use "{correct}" for OOXML connectors'
            )

    if issues:
        for issue in issues:
            err(f"{xml_path}: {issue}")
    else:
        ok(f"{xml_path}: transition/connector XML ok")


def check_nav_hyperlink_rid(zf, xml_path):
    """
    Navigation hyperlink actions (<a:hlinkClick action="ppaction://hlinkshowjump?...">)
    must carry r:id="" (empty string). Omitting r:id entirely is an OOXML schema
    violation — CT_Hyperlink requires the r:id attribute even when there is no relationship.
    Source: rep-x1 — PowerPoint added r:id="" to all navigation hlinkClick elements.
    """
    xml = zf.read(xml_path).decode("utf-8")
    if "ppaction://hlinkshowjump" not in xml:
        return

    # Find any hlinkClick with a ppaction but without r:id=
    for m in re.finditer(r'<a:hlinkClick\b([^>]*)>', xml):
        attrs = m.group(1)
        if "ppaction://hlinkshowjump" in attrs and 'r:id=' not in attrs:
            err(f"{xml_path}: <a:hlinkClick action=\"ppaction://...\"> missing r:id=\"\" "
                "(CT_Hyperlink requires r:id even for navigation actions)")
            return

    ok(f"{xml_path}: navigation hlinkClick r:id ok")


def check_bgpr_effectlst(zf, xml_path):
    """
    Any <p:bgPr> with a fill (blipFill, solidFill, or gradFill) must also include
    <a:effectLst/>. PowerPoint adds it on repair for all fill types.
    """
    xml = zf.read(xml_path).decode("utf-8")
    if "<p:bgPr>" not in xml:
        return
    # Find all <p:bgPr>...</p:bgPr> blocks
    for m in re.finditer(r'<p:bgPr>(.*?)</p:bgPr>', xml, re.DOTALL):
        block = m.group(1)
        has_fill = "blipFill" in block or "solidFill" in block or "gradFill" in block
        if has_fill and "effectLst" not in block:
            err(f"{xml_path}: <p:bgPr> with fill missing <a:effectLst/>")
            return
    ok(f"{xml_path}: <p:bgPr> effectLst ok")


def check_sptree_group_id(zf, xml_path):
    """
    The invisible root group shape's <p:cNvPr> inside <p:nvGrpSpPr> must have id="1".
    Using any other value (e.g. slide-offset IDs) causes repair.
    """
    xml = zf.read(xml_path).decode("utf-8")
    # Match <p:nvGrpSpPr><p:cNvPr id="N" .../>
    m = re.search(r'<p:nvGrpSpPr>\s*<p:cNvPr\s+id="(\d+)"', xml)
    if m:
        id_val = m.group(1)
        if id_val != "1":
            err(f"{xml_path}: spTree <p:nvGrpSpPr><p:cNvPr> id must be \"1\" (found \"{id_val}\")")
        else:
            ok(f"{xml_path}: spTree group cNvPr id=1 ok")


def check_content_types(zf):
    """
    [Content_Types].xml should not declare MIME type extensions for formats
    that aren't actually used in the package (causes warnings/repair).
    """
    ct_xml = zf.read("[Content_Types].xml").decode("utf-8")
    names = set(zf.namelist())

    # Find all Default extension declarations
    declared = re.findall(r'<Default Extension="([^"]+)"', ct_xml)
    used_exts = {n.rsplit(".", 1)[-1].lower() for n in names if "." in n}

    unused = [ext for ext in declared if ext.lower() not in used_exts and ext not in ("rels", "xml")]
    if unused:
        err(f"[Content_Types].xml: declares unused MIME extensions: {unused} — remove them")
    else:
        ok("[Content_Types].xml: no unused extension declarations")


# ─── Tier 1: Deep Semantic / Cross-Reference Checks ──────────────────────────

def _parse_rels(zf, rels_path):
    """Return {rId: (type, target, target_mode)} for a rels file."""
    if rels_path not in zf.namelist():
        return {}
    xml = zf.read(rels_path).decode("utf-8")
    root = ET.fromstring(xml)
    result = {}
    for rel in root:
        rid  = rel.get("Id", "")
        typ  = rel.get("Type", "")
        tgt  = rel.get("Target", "")
        mode = rel.get("TargetMode", "Internal")
        result[rid] = (typ, tgt, mode)
    return result


def _resolve_target(xml_part_path, target):
    """Resolve a relative rels Target to an absolute zip path."""
    # xml_part_path is like "ppt/slides/slide1.xml"
    # The rels file lives in "ppt/slides/_rels/slide1.xml.rels"
    # so the base for resolution is "ppt/slides/"
    base_dir = "/".join(xml_part_path.split("/")[:-1]) + "/"
    # Normalize ".." segments
    parts = (base_dir + target).split("/")
    resolved = []
    for p in parts:
        if p == "..":
            if resolved:
                resolved.pop()
        elif p and p != ".":
            resolved.append(p)
    return "/".join(resolved)


def check_relationship_integrity(zf, names):
    """
    Check 1: Orphaned rId — every r:id used in slide XML must exist in the rels file.
    Check 2: Missing physical asset — every rels Target (internal) must exist in ZIP.
    Check 3: Duplicate rId values within a single rels file.
    """
    slide_names = sorted(n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))

    for slide_path in slide_names:
        rels_path = slide_path.replace("slides/", "slides/_rels/") + ".rels"
        rels = _parse_rels(zf, rels_path)

        # Duplicate rIds
        all_rids = list(rels.keys())
        seen = set()
        for rid in all_rids:
            if rid in seen:
                err(f"{rels_path}: duplicate rId '{rid}'")
            seen.add(rid)

        # Collect all rId references in slide XML
        slide_xml = zf.read(slide_path).decode("utf-8")
        used_rids = set(re.findall(r'r:id="(rId\d+)"', slide_xml))
        # Also check embed= and link= attributes
        used_rids |= set(re.findall(r'r:embed="(rId\d+)"', slide_xml))
        used_rids |= set(re.findall(r'r:link="(rId\d+)"', slide_xml))

        for rid in used_rids:
            if rid not in rels:
                err(f"{slide_path}: uses {rid} which is not defined in {rels_path}")

        ok_count = 0
        for rid, (typ, target, mode) in rels.items():
            if mode == "External":
                ok_count += 1
                continue  # external hyperlinks don't need a physical file
            resolved = _resolve_target(slide_path, target)
            if resolved not in names:
                err(f"{rels_path}: {rid} → '{target}' resolved to '{resolved}' which does not exist in ZIP")
            else:
                ok_count += 1

        if ok_count > 0:
            ok(f"{rels_path}: {ok_count} relationship target(s) verified")


def check_root_rels(zf, names):
    """
    _rels/.rels: presentation.xml must be rId1.
    All Internal targets must exist in ZIP.
    """
    rels = _parse_rels(zf, "_rels/.rels")
    pres_rid = next((r for r, (t, _, _) in rels.items() if t == OFFICE_DOC_TYPE), None)
    if pres_rid != "rId1":
        err(f"_rels/.rels: presentation.xml (officeDocument) must be rId1 (found {pres_rid})")
    else:
        ok("_rels/.rels: presentation.xml=rId1")

    for rid, (typ, target, mode) in rels.items():
        if mode == "External":
            continue
        if target not in names:
            err(f"_rels/.rels: {rid} → '{target}' does not exist in ZIP")


def check_content_type_completeness(zf, names):
    """
    Every XML part in the ZIP must appear in [Content_Types].xml as an <Override>
    or be covered by a matching <Default> extension.
    The reverse: every <Override> path must exist in the ZIP.
    """
    if "[Content_Types].xml" not in names:
        err("[Content_Types].xml: file missing entirely")
        return

    ct_xml = zf.read("[Content_Types].xml").decode("utf-8")

    overrides = re.findall(r'<Override PartName="([^"]+)"', ct_xml)
    defaults  = re.findall(r'<Default Extension="([^"]+)"', ct_xml)
    default_exts = {e.lower() for e in defaults}

    # Check overrides reference existing parts
    for part in overrides:
        # PartName starts with "/" in [Content_Types], strip it
        part_clean = part.lstrip("/")
        if part_clean not in names:
            err(f"[Content_Types].xml: <Override PartName=\"{part}\"> references non-existent part")

    # Check all XML parts are registered
    xml_parts = [n for n in names if n.endswith(".xml") and not n.startswith("[")]
    override_parts = {p.lstrip("/") for p in overrides}
    for part in xml_parts:
        ext = part.rsplit(".", 1)[-1].lower()
        if part not in override_parts and ext not in default_exts:
            err(f"[Content_Types].xml: '{part}' is not registered (no <Override> or matching <Default>)")

    ok(f"[Content_Types].xml: Override/Default coverage checked ({len(overrides)} overrides, {len(default_exts)} default exts)")


def check_slide_layout_master_chain(zf, names):
    """
    For each slide:
      - slide → (rId1) → slideLayout must exist
      - slideLayout → (some rId) → slideMaster must exist
    All targets must be real files.
    """
    slide_names = sorted(n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))

    for slide_path in slide_names:
        rels_path = slide_path.replace("slides/", "slides/_rels/") + ".rels"
        slide_rels = _parse_rels(zf, rels_path)

        # Slide → slideLayout
        layout_entry = slide_rels.get("rId1")
        if not layout_entry:
            err(f"{slide_path}: rId1 (slideLayout) not found in rels")
            continue
        layout_target = _resolve_target(slide_path, layout_entry[1])
        if layout_target not in names:
            err(f"{slide_path}: slideLayout target '{layout_target}' not in ZIP")
            continue

        # slideLayout → slideMaster
        layout_rels_path = layout_target.replace("slideLayouts/", "slideLayouts/_rels/") + ".rels"
        layout_rels = _parse_rels(zf, layout_rels_path)
        master_entry = next(
            ((rid, e) for rid, e in layout_rels.items() if e[0] == SLIDE_MASTER_TYPE),
            None
        )
        if not master_entry:
            err(f"{layout_target}: no slideMaster relationship found")
            continue
        master_target = _resolve_target(layout_target, master_entry[1][1])
        if master_target not in names:
            err(f"{layout_target}: slideMaster target '{master_target}' not in ZIP")
        else:
            ok(f"{slide_path}: slide→layout→master chain ok")


def check_duplicate_cnvpr_ids(zf, names):
    """
    p:cNvPr id attributes must be unique within each slide (drawing part).
    Cross-slide uniqueness is NOT required by OOXML — only per-part uniqueness is.
    """
    slide_names = sorted(n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))

    for slide_path in slide_names:
        xml = zf.read(slide_path).decode("utf-8")
        ids = re.findall(r'<[a-z:]*cNvPr\b[^>]*\bid="(\d+)"', xml)
        seen = set()
        for i in ids:
            if i in seen:
                err(f"{slide_path}: duplicate cNvPr id={i} within slide")
            seen.add(i)
        if ids:
            ok(f"{slide_path}: cNvPr IDs unique within slide ({len(ids)} objects)")


def check_zero_byte_parts(zf, names):
    """
    Any XML part that is 0 bytes indicates a failed write stream.
    """
    bad = []
    for name in names:
        if name.endswith(".xml") or name.endswith(".rels"):
            info = zf.getinfo(name)
            if info.file_size == 0:
                bad.append(name)
    if bad:
        for b in bad:
            err(f"Zero-byte part: {b}")
    else:
        ok(f"Zero-byte check: all {sum(1 for n in names if n.endswith(('.xml','.rels')))} XML/rels parts non-empty")


def check_namespace_declarations(zf, names):
    """
    Every XML part must declare the namespaces it uses in its root element.
    Check for common prefixes (a:, p:, r:) used without declaration.
    """
    slide_names = sorted(n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))
    required_ns = {
        "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
        "p":  "http://schemas.openxmlformats.org/presentationml/2006/main",
        "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }

    for slide_path in slide_names:
        xml = zf.read(slide_path).decode("utf-8")
        # Skip the XML declaration (<?xml ... ?>), then match the root element
        stripped = re.sub(r'^\s*<\?[^?]*\?>\s*', '', xml)
        root_tag_match = re.match(r'<[^>]+>', stripped, re.DOTALL)
        if not root_tag_match:
            err(f"{slide_path}: cannot find root element")
            continue
        root_tag = root_tag_match.group(0)

        for prefix, uri in required_ns.items():
            used = bool(re.search(rf'\b{re.escape(prefix)}:', xml))
            declared = uri in root_tag
            if used and not declared:
                err(f"{slide_path}: namespace '{prefix}:' is used but not declared (missing xmlns:{prefix})")

    ok(f"Namespace declarations checked for {len(slide_names)} slide(s)")


def check_notes_master_has_own_theme(zf, names):
    """
    The notes master must reference its own dedicated theme file (not the same
    theme1.xml used by the slide master). PowerPoint adds theme2.xml on repair
    if this is missing.
    """
    nm_rels = "ppt/notesMasters/_rels/notesMaster1.xml.rels"
    if nm_rels not in names:
        err(f"{nm_rels}: missing")
        return
    rels = _parse_rels(zf, nm_rels)
    THEME_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    theme_entry = next(((rid, e) for rid, e in rels.items() if e[0] == THEME_TYPE), None)
    if not theme_entry:
        err(f"{nm_rels}: no theme relationship found")
        return
    rid, (typ, target, mode) = theme_entry
    # Resolve and check the target exists
    resolved = _resolve_target("ppt/notesMasters/notesMaster1.xml", target)
    if resolved not in names:
        err(f"{nm_rels}: theme target '{resolved}' does not exist in ZIP")
        return
    # Check it's NOT the same file as used by slideMaster
    sm_rels = "ppt/slideMasters/_rels/slideMaster1.xml.rels"
    sm_rels_data = _parse_rels(zf, sm_rels)
    sm_theme = next(
        (_resolve_target("ppt/slideMasters/slideMaster1.xml", e[1])
         for _, e in sm_rels_data.items() if e[0] == THEME_TYPE),
        None
    )
    if sm_theme and resolved == sm_theme:
        err(f"{nm_rels}: notes master shares theme with slide master ({resolved}) — needs dedicated theme file")
    else:
        ok(f"{nm_rels}: notes master has dedicated theme ({resolved})")


def check_duplicate_rel_targets(zf, names):
    """
    Multiple rIds in one rels file pointing to the same target (non-hyperlink)
    can confuse PowerPoint's internal cache.
    """
    all_rels = sorted(n for n in names if n.endswith(".rels") and "_rels/" in n)
    for rels_path in all_rels:
        rels = _parse_rels(zf, rels_path)
        target_counts = {}
        for rid, (typ, target, mode) in rels.items():
            if mode == "External":
                continue  # multiple rIds for same URL is fine (hyperlinks)
            key = (typ, target)
            target_counts.setdefault(key, []).append(rid)
        for (typ, target), rids in target_counts.items():
            if len(rids) > 1:
                # Image deduplication legitimately creates multiple rIds for the same media file
                if "/relationships/image" in typ or "/relationships/video" in typ or "/relationships/audio" in typ:
                    continue
                err(f"{rels_path}: target '{target}' referenced by multiple rIds: {rids}")

    ok(f"Duplicate rels targets: checked {len(all_rels)} rels file(s)")


def check_presentation_slide_rids(zf, names):
    """
    Verify that slide rIds in presentation.xml match the rIds in
    presentation.xml.rels. A mismatch (e.g. after slide removal) causes
    PowerPoint to trigger the repair dialog.
    """
    pres_path = "ppt/presentation.xml"
    rels_path = "ppt/_rels/presentation.xml.rels"
    if pres_path not in names or rels_path not in names:
        return

    # Get slide rIds from presentation.xml sldIdLst
    tree = ET.parse(zf.open(pres_path))
    root = tree.getroot()
    sld_ids = root.findall(".//p:sldIdLst/p:sldId", NS)
    pres_rids = {}
    for el in sld_ids:
        rid = el.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        sid = el.get("id")
        if rid:
            pres_rids[rid] = sid

    # Get slide rIds from presentation.xml.rels
    rels = _parse_rels(zf, rels_path)
    rels_slide_rids = {}
    for rid, (typ, target, _mode) in rels.items():
        if "/relationships/slide" in typ and "slideMaster" not in target and "slideLayout" not in target:
            rels_slide_rids[rid] = target

    # Each rId in presentation.xml must exist in the rels
    for rid, sid in pres_rids.items():
        if rid not in rels_slide_rids:
            err(f"{pres_path}: sldId id=\"{sid}\" references {rid} which is not a slide relationship in {rels_path}")

    ok(f"Presentation slide rIds: {len(pres_rids)} slide(s) verified")


VALID_UNDERLINE_VALUES = {
    "none", "words", "sng", "dbl", "heavy",
    "dotted", "dottedHeavy",
    "dash", "dashHeavy", "dashLong", "dashLongHeavy",
    "dotDash", "dotDashHeavy",
    "wavy", "wavyHeavy", "wavyDbl",
}


def check_underline_values(zf, names):
    """
    Validate that all u= attributes on <a:rPr> use valid ST_TextUnderlineType values.
    Invalid values like "heavyDash" (should be "dashHeavy") cause PowerPoint to strip
    the underline during repair.
    """
    slide_paths = sorted(n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml"))
    checked = 0
    for path in slide_paths:
        xml = zf.read(path).decode("utf-8", errors="replace")
        for m in re.finditer(r'<a:rPr\b[^>]*\bu="([^"]+)"', xml):
            val = m.group(1)
            checked += 1
            if val not in VALID_UNDERLINE_VALUES:
                err(f"{path}: invalid underline value u=\"{val}\" — not a valid ST_TextUnderlineType")
    ok(f"Underline values: checked {checked} attribute(s) across {len(slide_paths)} slide(s)")


def check_chart_xml(zf, chart_path):
    """Catch invalid chart XML patterns that trigger PowerPoint repair."""
    try:
        data = zf.read(chart_path)
    except Exception:
        err(f"{chart_path}: cannot read")
        return

    text = data.decode("utf-8", errors="replace")

    # <c:tickMark> is not a valid OOXML element — the schema requires
    # <c:majorTickMark> and <c:minorTickMark> instead.
    if b"<c:tickMark " in data or b"<c:tickMark>" in data:
        err(f"{chart_path}: invalid <c:tickMark> element — use <c:majorTickMark> and <c:minorTickMark>")
        return

    # <a:noFill/> must not appear directly before <a:ln> inside a series <c:spPr>.
    # The chart-space-level <c:spPr><a:noFill/><a:ln> is valid and must NOT be flagged.
    for ser_block in re.findall(r"<c:ser\b.*?</c:ser>", text, re.DOTALL):
        if re.search(r"<c:spPr>\s*<a:noFill/>\s*<a:ln\b", ser_block):
            err(f"{chart_path}: <a:noFill/> before <a:ln> inside series <c:spPr> — PowerPoint removes this during repair")
            return

    # Every chart element must have a chart-level <c:dLbls> block after its
    # series and before <c:axId>/<c:firstSliceAng>. Its absence triggers repair.
    chart_el_tags = ["barChart", "lineChart", "pieChart", "doughnutChart", "areaChart", "scatterChart"]
    for tag in chart_el_tags:
        if f"<c:{tag}>" in text or f"<c:{tag} " in text:
            # Extract the chart element body
            m = re.search(rf"<c:{tag}[\s>].*?</c:{tag}>", text, re.DOTALL)
            if m and "<c:dLbls>" not in m.group():
                err(f"{chart_path}: missing chart-level <c:dLbls> inside <c:{tag}>")
                return

    # Horizontal bar chart (<c:barDir val="bar"/>) must have catAx on the left
    # (axPos="l") and valAx on the bottom (axPos="b"), not the reverse.
    if '<c:barDir val="bar"/>' in text:
        m = re.search(r"<c:catAx>.*?</c:catAx>", text, re.DOTALL)
        if m and '<c:axPos val="b"/>' in m.group():
            err(f"{chart_path}: horizontal bar chart has catAx axPos=\"b\" — should be \"l\"")
            return

    # Per-series <c:dLbls> must come before <c:cat>/<c:xVal> per OOXML content
    # model (CT_BarSer, CT_LineSer, etc.). Emitting dLbls after cat/val triggers repair.
    for ser_block in re.findall(r"<c:ser\b.*?</c:ser>", text, re.DOTALL):
        cat_pos  = ser_block.find("<c:cat>")
        xval_pos = ser_block.find("<c:xVal>")
        dlbls_pos = ser_block.find("<c:dLbls>")
        if dlbls_pos == -1:
            continue
        ref_pos = min(p for p in [cat_pos, xval_pos] if p != -1) if any(p != -1 for p in [cat_pos, xval_pos]) else -1
        if ref_pos != -1 and dlbls_pos > ref_pos:
            err(f"{chart_path}: series <c:dLbls> appears after <c:cat>/<c:xVal> — must come before per OOXML content model")
            return

    ok(f"{chart_path}: chart XML ok")


def main():
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <file.pptx>")
        sys.exit(1)

    pptx_path = sys.argv[1]
    print(f"\nChecking repair patterns in: {pptx_path}\n")

    try:
        zf = zipfile.ZipFile(pptx_path)
    except Exception as e:
        print(f"ERROR: Cannot open {pptx_path}: {e}")
        sys.exit(1)

    names = set(zf.namelist())

    # ── Tier 0: PowerPoint semantic conventions ──────────────

    # Check slide rels ordering (rId1=layout, rId2=notes)
    slide_names = sorted(n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n))
    for slide in slide_names:
        check_slide_rels(zf, slide)

    print()

    # Check notes slides for empty runs
    notes_names = sorted(n for n in names if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n))
    for notes in notes_names:
        check_notes_slide(zf, notes)

    print()

    # Check slide XML for empty open/close elements, shadow defaults, and animation patterns
    for slide in slide_names:
        check_empty_elements(zf, slide)
        check_shadow_defaults(zf, slide)
        check_bgpr_effectlst(zf, slide)
        check_sptree_group_id(zf, slide)
        check_animation_xml(zf, slide)
        check_transition_xml(zf, slide)
        check_nav_hyperlink_rid(zf, slide)

    print()

    # Check presentation.xml section namespace placement
    check_presentation_sections(zf)
    print()

    # Check chart XML for invalid patterns
    chart_names = sorted(n for n in names if re.match(r"ppt/charts/chart\d+\.xml$", n))
    for chart in chart_names:
        check_chart_xml(zf, chart)
    if chart_names:
        print()

    # Check content types for unused extension declarations
    if "[Content_Types].xml" in names:
        check_content_types(zf)

    print()

    # ── Tier 1: Deep cross-reference / semantic integrity ─────

    print("=== Deep Semantic Checks ===")
    print()

    # Root rels: presentation.xml must be rId1
    check_root_rels(zf, names)
    print()

    # Cross-reference rIds in slide XML with rels files and physical files
    check_relationship_integrity(zf, names)
    print()

    # All XML parts must be registered in [Content_Types].xml
    check_content_type_completeness(zf, names)
    print()

    # Slide → slideLayout → slideMaster chain must be consistent
    check_slide_layout_master_chain(zf, names)
    print()

    # p:cNvPr id attributes must be unique within each slide
    check_duplicate_cnvpr_ids(zf, names)
    print()

    # Notes master must have its own dedicated theme file
    check_notes_master_has_own_theme(zf, names)
    print()

    # No zero-byte XML or rels parts
    check_zero_byte_parts(zf, names)
    print()

    # Required namespace prefixes must be declared
    check_namespace_declarations(zf, names)
    print()

    # No duplicate rels targets within a single rels file
    check_duplicate_rel_targets(zf, names)
    print()

    # Slide rIds in presentation.xml must match presentation.xml.rels
    check_presentation_slide_rids(zf, names)
    print()

    # Invalid underline style values
    check_underline_values(zf, names)
    print()

    zf.close()

    if errors:
        print(f"RESULT: {len(errors)} repair-pattern issue(s) found:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)
    else:
        print("RESULT: No repair-pattern issues found.")
        sys.exit(0)


if __name__ == "__main__":
    main()
