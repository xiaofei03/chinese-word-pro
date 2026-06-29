from __future__ import annotations

import argparse
import re
import shutil
import tempfile
import zipfile
from pathlib import Path


BLUE_HEX_VALUES = ["4F81BD", "365F91", "0563C1", "0000FF"]
AUDIT_PATTERNS = [
    "4F81BD",
    "365F91",
    "0563C1",
    "0000FF",
    'w:themeColor="accent1"',
    'w:themeFill="accent1"',
    "<w:pBdr>",
    "<w:u ",
]


def read_zip_text(zf: zipfile.ZipFile, member: str) -> str:
    return zf.read(member).decode("utf-8", errors="ignore")


def clean_styles_xml(xml: str) -> str:
    xml = re.sub(r"<w:pBdr>.*?</w:pBdr>", "", xml, flags=re.S)
    xml = re.sub(r"<w:u\b[^>]*/>", "", xml)
    xml = re.sub(r'w:themeColor="accent1"', 'w:themeColor="text1"', xml)
    xml = re.sub(r'w:themeFill="accent1"', 'w:themeFill="text1"', xml)
    for hex_value in BLUE_HEX_VALUES:
        xml = xml.replace(hex_value, "000000")
    return xml


def clean_document_xml(xml: str) -> str:
    xml = re.sub(r"<w:u\b[^>]*/>", "", xml)
    for hex_value in BLUE_HEX_VALUES:
        xml = xml.replace(hex_value, "000000")
    xml = re.sub(r'w:themeColor="accent1"', 'w:themeColor="text1"', xml)
    xml = re.sub(r'w:themeFill="accent1"', 'w:themeFill="text1"', xml)
    return xml


def clean_theme_xml(xml: str) -> str:
    xml = re.sub(
        r"<a:accent1>.*?</a:accent1>",
        '<a:accent1><a:srgbClr val="000000"/></a:accent1>',
        xml,
        flags=re.S,
    )
    xml = re.sub(
        r"<a:hlink>.*?</a:hlink>",
        '<a:hlink><a:srgbClr val="000000"/></a:hlink>',
        xml,
        flags=re.S,
    )
    xml = re.sub(
        r"<a:folHlink>.*?</a:folHlink>",
        '<a:folHlink><a:srgbClr val="000000"/></a:folHlink>',
        xml,
        flags=re.S,
    )
    for hex_value in BLUE_HEX_VALUES:
        xml = xml.replace(hex_value, "000000")
    return xml


def audit_docx(path: Path) -> dict[str, list[str]]:
    result: dict[str, list[str]] = {}
    with zipfile.ZipFile(path) as zf:
        for member in ["word/document.xml", "word/styles.xml", "word/theme/theme1.xml"]:
            xml = read_zip_text(zf, member)
            hits = [pat for pat in AUDIT_PATTERNS if pat in xml]
            result[member] = hits
    return result


def rewrite_docx(input_docx: Path, output_docx: Path) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        with zipfile.ZipFile(input_docx) as zf:
            zf.extractall(tmpdir)

        document = tmpdir / "word/document.xml"
        styles = tmpdir / "word/styles.xml"
        theme = tmpdir / "word/theme/theme1.xml"

        if styles.exists():
            styles.write_text(clean_styles_xml(styles.read_text(encoding="utf-8")), encoding="utf-8")
        if document.exists():
            document.write_text(clean_document_xml(document.read_text(encoding="utf-8")), encoding="utf-8")
        if theme.exists():
            theme.write_text(clean_theme_xml(theme.read_text(encoding="utf-8")), encoding="utf-8")

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zf:
            for file in tmpdir.rglob("*"):
                if file.is_file():
                    zf.write(file, file.relative_to(tmpdir))


def format_audit(audit: dict[str, list[str]]) -> str:
    lines = []
    for member, hits in audit.items():
        lines.append(f"{member}: {hits if hits else 'OK'}")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Clean Word theme/style blue residue from Chinese report DOCX files and audit the result."
    )
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    parser.add_argument("--in-place", action="store_true")
    parser.add_argument("--keep-backup", action="store_true")
    args = parser.parse_args()

    input_docx = Path(args.input_docx).resolve()
    output_docx = Path(args.output_docx).resolve()

    rewrite_docx(input_docx, output_docx)
    audit = audit_docx(output_docx)
    blocking = {
        member: hits
        for member, hits in audit.items()
        if any(hit in hits for hit in ['4F81BD', '365F91', '0563C1', '0000FF', 'w:themeColor="accent1"', 'w:themeFill="accent1"', '<w:pBdr>'])
    }
    if blocking:
        raise SystemExit("Theme cleanup failed audit:\n" + format_audit(audit))

    if args.in_place and output_docx != input_docx:
        backup = input_docx.with_suffix(".pre-theme-clean.docx")
        if args.keep_backup:
            shutil.copy2(input_docx, backup)
        output_docx.replace(input_docx)
        print(f"OK: cleaned in place -> {input_docx}")
    else:
        print(f"OK: {output_docx}")
    print(format_audit(audit))


if __name__ == "__main__":
    main()
