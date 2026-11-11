#!/usr/bin/env python3
"""arabic-first-docs — build, check and repair Arabic-first Word documents.

The problem this solves: python-docx makes it easy to set `w:bidi` in the wrong
place inside `w:pPr`. Word then ignores it silently and the document opens
left-to-right, which is discovered by the reader, not the author. The fix is
structural, at three levels, and it is checkable, so this tool checks it.

Commands:
    build SPEC.json OUT.docx [--lang ar|en]
    check OUT.docx [--json]
    fix IN.docx OUT.docx

Exit codes: 0 pass, 1 the document fails the RTL checks, 2 usage or input error.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ARABIC = re.compile(r"[\u0600-\u06FF]")
ARABIC_FONT = "Arial"


# --------------------------------------------------------------------------
# the three-level fix
# --------------------------------------------------------------------------
#
# Word ignores an element that sits out of schema order, silently. So every
# insertion here goes through the schema sequence, not append().

RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
             "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
             "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
             "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
             "eastAsianLayout", "specVanish", "oMath"]

PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
             "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
             "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
             "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
             "suppressOverlap", "jc", "textDirection", "textAlignment",
             "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]

SECTPR_ORDER = ["footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc",
                "pgBorders", "lnNumType", "pgNumType", "cols", "formProt",
                "vAlign", "noEndnote", "titlePg", "textDirection", "bidi",
                "rtlGutter", "docGrid", "printerSettings", "sectPrChange"]

TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
               "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
               "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
               "tblCellMar", "tblLook", "tblCaption", "tblDescription"]


def insert_ordered(parent, element, order: list[str]):
    """Insert element at its schema position, or append when unknown."""
    name = element.tag.split("}")[-1]
    if name not in order:
        parent.append(element)
        return element
    index = order.index(name)
    for child in parent:
        child_name = child.tag.split("}")[-1]
        if child_name in order and order.index(child_name) > index:
            child.addprevious(element)
            return element
    parent.append(element)
    return element


def apply_doc_defaults(doc) -> None:
    """Level 1: RTL at the style defaults, so every element inherits it."""
    styles_el = doc.styles.element
    defaults = styles_el.find(qn("w:docDefaults"))
    if defaults is None:
        defaults = OxmlElement("w:docDefaults")
        styles_el.insert(0, defaults)

    rpr_default = defaults.find(qn("w:rPrDefault"))
    if rpr_default is None:
        rpr_default = OxmlElement("w:rPrDefault")
        defaults.insert(0, rpr_default)
    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        rpr_default.append(rpr)
    if rpr.find(qn("w:rtl")) is None:
        insert_ordered(rpr, OxmlElement("w:rtl"), RPR_ORDER)

    ppr_default = defaults.find(qn("w:pPrDefault"))
    if ppr_default is None:
        ppr_default = OxmlElement("w:pPrDefault")
        defaults.append(ppr_default)
    ppr = ppr_default.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        ppr_default.append(ppr)
    if ppr.find(qn("w:bidi")) is None:
        insert_ordered(ppr, OxmlElement("w:bidi"), PPR_ORDER)


def paragraph_bidi(paragraph) -> None:
    """Level 2: w:bidi in valid schema position, before spacing/indent/jc."""
    ppr = paragraph._p.get_or_add_pPr()
    if ppr.find(qn("w:bidi")) is None:
        insert_ordered(ppr, OxmlElement("w:bidi"), PPR_ORDER)


def run_rtl(run, font: str = ARABIC_FONT) -> None:
    """Level 3: w:rtl on the run plus the complex-script font, or Arabic falls back."""
    rpr = run._element.get_or_add_rPr()
    if rpr.find(qn("w:rtl")) is None:
        insert_ordered(rpr, OxmlElement("w:rtl"), RPR_ORDER)
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = insert_ordered(rpr, OxmlElement("w:rFonts"), RPR_ORDER)
    fonts.set(qn("w:ascii"), font)
    fonts.set(qn("w:hAnsi"), font)
    fonts.set(qn("w:cs"), font)


def section_bidi(section) -> None:
    sect = section._sectPr
    if sect.find(qn("w:bidi")) is None:
        insert_ordered(sect, OxmlElement("w:bidi"), SECTPR_ORDER)


def table_bidi(table) -> None:
    tblpr = table._tbl.tblPr
    if tblpr.find(qn("w:bidiVisual")) is None:
        insert_ordered(tblpr, OxmlElement("w:bidiVisual"), TBLPR_ORDER)
    if tblpr.find(qn("w:tblLayout")) is None:
        layout = OxmlElement("w:tblLayout")
        layout.set(qn("w:type"), "fixed")
        insert_ordered(tblpr, layout, TBLPR_ORDER)


# --------------------------------------------------------------------------
# build
# --------------------------------------------------------------------------

def add_paragraph(doc, text: str, *, bold=False, size=None, align=None):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    run.bold = bold
    if size:
        run.font.size = Pt(size)
    # In a bidi paragraph w:jc="right" means the END side, which is the left
    # edge visually. Leaving alignment unset gives Word's "start", i.e. rtl.
    if align is not None:
        paragraph.alignment = align
    paragraph_bidi(paragraph)
    run_rtl(run)
    fmt = paragraph.paragraph_format
    fmt.space_after = Pt(6)
    fmt.line_spacing = 1.15
    if size and size >= 14:
        fmt.space_before = Pt(14)
    return paragraph


def build(spec: dict, out: str, lang: str = "ar") -> dict:
    doc = Document()

    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, attr, Cm(2.5))

    if lang == "ar":
        apply_doc_defaults(doc)
        section_bidi(section)

    title = spec.get("title")
    if title:
        add_paragraph(doc, title, bold=True, size=18)
    subtitle = spec.get("subtitle")
    if subtitle:
        add_paragraph(doc, subtitle, size=12)

    for block in spec.get("sections", []):
        heading = block.get("heading")
        if heading:
            add_paragraph(doc, heading, bold=True, size=14)
        for paragraph in block.get("paragraphs", []):
            add_paragraph(doc, paragraph)
        rows = block.get("table")
        if rows:
            width = len(rows[0])
            table = doc.add_table(rows=len(rows), cols=width)
            table.style = "Table Grid"
            for r_index, row in enumerate(rows):
                for c_index, value in enumerate(row):
                    cell = table.cell(r_index, c_index)
                    cell.text = ""
                    paragraph = cell.paragraphs[0]
                    run = paragraph.add_run(str(value))
                    if r_index == 0:
                        run.bold = True
                    paragraph_bidi(paragraph)
                    run_rtl(run)
            total = Cm(16)
            for column in table.columns:
                column_width = int(total / max(len(table.columns), 1))
                for cell in column.cells:
                    cell.width = column_width
            if lang == "ar":
                table_bidi(table)

    doc.save(out)
    return inspect_document(out)


# --------------------------------------------------------------------------
# check
# --------------------------------------------------------------------------

@dataclass
class Finding:
    code: str
    detail: str


def inspect_document(path: str) -> list[Finding]:
    """Count what the eye cannot see: bidi, rtl and complex-script fonts."""
    doc = Document(path)
    findings: list[Finding] = []

    defaults = doc.styles.element.find(qn("w:docDefaults"))
    has_bidi_default = (
        defaults is not None
        and defaults.find(f"{qn('w:pPrDefault')}/{qn('w:pPr')}/{qn('w:bidi')}") is not None
    )
    if not has_bidi_default:
        missing_defaults = True
    else:
        missing_defaults = False

    missing_sections: list[int] = []

    arabic_paragraphs = 0
    missing_para_bidi = 0
    missing_run_rtl = 0
    missing_cs_font = 0

    for paragraph in doc.paragraphs:
        if not ARABIC.search(paragraph.text or ""):
            continue
        arabic_paragraphs += 1
        ppr = paragraph._p.pPr
        if ppr is None or ppr.find(qn("w:bidi")) is None:
            missing_para_bidi += 1
        for run in paragraph.runs:
            rpr = run._element.rPr
            if rpr is None or rpr.find(qn("w:rtl")) is None:
                missing_run_rtl += 1
                continue
            fonts = rpr.find(qn("w:rFonts"))
            if fonts is None or not fonts.get(qn("w:cs")):
                missing_cs_font += 1

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if ARABIC.search(paragraph.text or ""):
                        arabic_paragraphs += 1
                        ppr = paragraph._p.pPr
                        if ppr is None or ppr.find(qn("w:bidi")) is None:
                            missing_para_bidi += 1
        if table._tbl.tblPr.find(qn("w:bidiVisual")) is None:
            findings.append(Finding("NO_TABLE_BIDI_VISUAL",
                                    "a table reads left-to-right inside an Arabic document"))

    for index, section in enumerate(doc.sections, start=1):
        if section._sectPr.find(qn("w:bidi")) is None:
            missing_sections.append(index)

    if arabic_paragraphs and missing_defaults:
        findings.append(Finding("NO_DOC_DEFAULTS",
                                "document style defaults carry no w:bidi"))
    for index in missing_sections if arabic_paragraphs else []:
        findings.append(Finding("NO_SECTION_BIDI",
                                f"section {index} carries no w:bidi"))

    if arabic_paragraphs and missing_para_bidi:
        findings.append(Finding("MISSING_PARA_BIDI",
                                f"{missing_para_bidi} of {arabic_paragraphs} Arabic paragraphs carry no w:bidi"))
    if missing_run_rtl:
        findings.append(Finding("MISSING_RUN_RTL",
                                f"{missing_run_rtl} runs carry no w:rtl"))
    if missing_cs_font:
        findings.append(Finding("MISSING_CS_FONT",
                                f"{missing_cs_font} runs set no complex-script font"))
    return findings


def check(path: str) -> list[Finding]:
    try:
        return inspect_document(path)
    except Exception as exc:  # noqa: BLE001 - a bad file is a finding, not a crash
        return [Finding("UNREADABLE", f"cannot open the document: {exc}")]


# --------------------------------------------------------------------------
# fix
# --------------------------------------------------------------------------

def fix(source: str, out: str) -> list[Finding]:
    """Repair an existing document: the same three levels, applied in place."""
    doc = Document(source)
    apply_doc_defaults(doc)
    for paragraph in doc.paragraphs:
        if ARABIC.search(paragraph.text or ""):
            paragraph_bidi(paragraph)
            ppr = paragraph._p.pPr
            jc = ppr.find(qn("w:jc")) if ppr is not None else None
            if jc is not None and jc.get(qn("w:val")) in ("left", "right"):
                ppr.remove(jc)
            for run in paragraph.runs:
                run_rtl(run)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if ARABIC.search(paragraph.text or ""):
                        paragraph_bidi(paragraph)
                        for run in paragraph.runs:
                            run_rtl(run)
        table_bidi(table)
    for section in doc.sections:
        section_bidi(section)
    doc.save(out)
    return inspect_document(out)


# --------------------------------------------------------------------------
# cli
# --------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Arabic-first Word documents: build, check, fix.")
    sub = parser.add_subparsers(dest="command", required=True)

    p_build = sub.add_parser("build", help="build a document from a JSON spec")
    p_build.add_argument("spec")
    p_build.add_argument("out")
    p_build.add_argument("--lang", default="ar", choices=["ar", "en"])

    p_check = sub.add_parser("check", help="verify the RTL properties of a document")
    p_check.add_argument("docx")
    p_check.add_argument("--json", action="store_true")

    p_fix = sub.add_parser("fix", help="repair a document that opens left-to-right")
    p_fix.add_argument("source")
    p_fix.add_argument("out")

    args = parser.parse_args(argv)

    if args.command == "build":
        try:
            with open(args.spec, encoding="utf-8") as fh:
                spec = json.load(fh)
        except (FileNotFoundError, json.JSONDecodeError) as exc:
            print(f"error: cannot read spec: {exc}", file=sys.stderr)
            return 2
        findings = build(spec, args.out, args.lang)
        target = args.out
    elif args.command == "fix":
        findings = fix(args.source, args.out)
        target = args.out
    else:
        findings = check(args.docx)
        if args.json:
            print(json.dumps({"document": args.docx,
                              "findings": [f.__dict__ for f in findings],
                              "result": "PASS" if not findings else "FAIL"}, indent=2))
            return 0 if not findings else 1
        target = args.docx

    print(f"arabic-first-docs  ·  {target}")
    if findings:
        for finding in findings:
            print(f"  {finding.code:<24} {finding.detail}")
        print(f"FAIL: {len(findings)} finding{'' if len(findings) == 1 else 's'}")
        return 1
    print("PASS: docDefaults bidi, paragraph bidi, run rtl, complex-script font, section and table direction")
    return 0


if __name__ == "__main__":
    sys.exit(main())
