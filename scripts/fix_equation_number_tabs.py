#!/usr/bin/env python3
"""Repair right-aligned numbering for native Word equations.

The script preserves native OMML equation objects. It rewrites tab stops for
numbered formula paragraphs and handles short continuation-tail rows by using a
right tab stop only, preventing equation numbers from stopping at a center tab.
"""

from __future__ import annotations

import argparse
import re
import shutil
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}
NUMBER_RE = re.compile(r"^\s*[(（]\d+[)）]\s*$")


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def text_of(el: etree._Element) -> str:
    return "".join(el.xpath(".//w:t/text()", namespaces=NS))


def math_text_of(el: etree._Element) -> str:
    return "".join(el.xpath(".//m:t/text()", namespaces=NS))


def has_math(el: etree._Element) -> bool:
    return (
        el.tag in {qn("m", "oMath"), qn("m", "oMathPara")}
        or bool(el.xpath(".//m:oMath|.//m:oMathPara", namespaces=NS))
    )


def ensure_child(parent: etree._Element, name: str) -> etree._Element:
    child = parent.find(qn("w", name))
    if child is None:
        child = etree.Element(qn("w", name))
        parent.append(child)
    return child


def ensure_ppr(p: etree._Element) -> etree._Element:
    ppr = p.find(qn("w", "pPr"))
    if ppr is None:
        ppr = etree.Element(qn("w", "pPr"))
        p.insert(0, ppr)
    return ppr


def set_formula_tabs(p: etree._Element, *, center: int | None = 4500, right: int = 9000) -> None:
    ppr = ensure_ppr(p)
    jc = ensure_child(ppr, "jc")
    jc.set(qn("w", "val"), "left")

    ind = ensure_child(ppr, "ind")
    for key in ["left", "right", "firstLine", "firstLineChars"]:
        ind.set(qn("w", key), "0")
    for key in ["hanging", "hangingChars"]:
        ind.attrib.pop(qn("w", key), None)

    spacing = ensure_child(ppr, "spacing")
    spacing.set(qn("w", "before"), "120")
    spacing.set(qn("w", "after"), "120")
    spacing.set(qn("w", "lineRule"), "auto")
    spacing.set(qn("w", "line"), "240")

    tabs = ppr.find(qn("w", "tabs"))
    if tabs is not None:
        ppr.remove(tabs)
    tabs = etree.Element(qn("w", "tabs"))
    if center is not None:
        center_tab = etree.Element(qn("w", "tab"))
        center_tab.set(qn("w", "val"), "center")
        center_tab.set(qn("w", "pos"), str(center))
        tabs.append(center_tab)
    right_tab = etree.Element(qn("w", "tab"))
    right_tab.set(qn("w", "val"), "right")
    right_tab.set(qn("w", "pos"), str(right))
    tabs.append(right_tab)
    ppr.append(tabs)


def make_tab_run() -> etree._Element:
    run = etree.Element(qn("w", "r"))
    run.append(etree.Element(qn("w", "tab")))
    return run


def make_number_run(number: str) -> etree._Element:
    run = etree.Element(qn("w", "r"))
    rpr = etree.SubElement(run, qn("w", "rPr"))
    fonts = etree.SubElement(rpr, qn("w", "rFonts"))
    fonts.set(qn("w", "ascii"), "Times New Roman")
    fonts.set(qn("w", "hAnsi"), "Times New Roman")
    size = etree.SubElement(rpr, qn("w", "sz"))
    size.set(qn("w", "val"), "22")
    text = etree.SubElement(run, qn("w", "t"))
    text.text = number
    return run


def child_is_tab_run(child: etree._Element) -> bool:
    return child.tag == qn("w", "r") and child.find(qn("w", "tab")) is not None


def child_is_number_run(child: etree._Element) -> bool:
    return child.tag == qn("w", "r") and NUMBER_RE.match(text_of(child).strip() or "") is not None


def cleanup_formula_paragraph(p: etree._Element) -> str | None:
    number = None
    for child in list(p):
        if child_is_number_run(child):
            number = text_of(child).strip()
            p.remove(child)
        elif child_is_tab_run(child):
            p.remove(child)
    return number


def set_math_size(el: etree._Element, half_points: int) -> None:
    for r in el.xpath(".//m:r", namespaces=NS):
        wrpr = r.find(qn("w", "rPr"))
        if wrpr is None:
            wrpr = etree.Element(qn("w", "rPr"))
            mrpr = r.find(qn("m", "rPr"))
            if mrpr is not None:
                mrpr.addnext(wrpr)
            else:
                r.insert(0, wrpr)
        for tag in ["sz", "szCs"]:
            node = wrpr.find(qn("w", tag))
            if node is None:
                node = etree.SubElement(wrpr, qn("w", tag))
            node.set(qn("w", "val"), str(half_points))


def add_number_layout(p: etree._Element, number: str, *, short_tail_line: bool) -> None:
    set_formula_tabs(p, center=None if short_tail_line else 4500)
    if not short_tail_line:
        first_math = next((child for child in p if has_math(child)), None)
        if first_math is not None:
            first_math.addprevious(make_tab_run())
    p.append(make_tab_run())
    p.append(make_number_run(number))


def repair_numbered_formula(p: etree._Element) -> str | None:
    if not has_math(p):
        return None
    number = cleanup_formula_paragraph(p)
    if not number:
        return None

    math_text = math_text_of(p)
    formula_len = len(math_text)
    short_tail_line = formula_len < 35 and math_text.lstrip().startswith("+")

    if formula_len > 75:
        set_math_size(p, 17)
    elif formula_len > 55:
        set_math_size(p, 18)

    add_number_layout(p, number, short_tail_line=short_tail_line)
    return number


def fix_document_xml(xml: bytes) -> tuple[bytes, list[str]]:
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.fromstring(xml, parser)
    fixed: list[str] = []
    for p in root.xpath("//w:body/w:p", namespaces=NS):
        number = repair_numbered_formula(p)
        if number:
            fixed.append(number)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), fixed


def patch_docx(input_docx: Path, output_docx: Path) -> list[str]:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        with zipfile.ZipFile(input_docx) as zin:
            zin.extractall(tmp_path)
        document_xml = tmp_path / "word" / "document.xml"
        patched, fixed = fix_document_xml(document_xml.read_bytes())
        document_xml.write_bytes(patched)
        with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for file in tmp_path.rglob("*"):
                if file.is_file():
                    zout.write(file, file.relative_to(tmp_path).as_posix())
    shutil.copystat(input_docx, output_docx, follow_symlinks=True)
    return fixed


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-docx", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    args = parser.parse_args()

    fixed = patch_docx(args.input_docx, args.output_docx)
    print(f"fixed {len(fixed)} numbered formulas: {', '.join(fixed)}")


if __name__ == "__main__":
    main()
