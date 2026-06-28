from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path


FONT_REPLACEMENTS = {
    "宋体": "宋体-简",
    "黑体": "黑体-简",
}


def replace_fonts_in_xml(xml: str) -> str:
    for source, target in FONT_REPLACEMENTS.items():
        xml = xml.replace(source, target)
    return xml


def make_pdf_safe_docx(input_docx: Path, output_docx: Path) -> None:
    if not input_docx.exists():
        raise FileNotFoundError(input_docx)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        with zipfile.ZipFile(input_docx) as src_zip:
            src_zip.extractall(tmp_path)

        for rel in ["word/document.xml", "word/styles.xml", "word/fontTable.xml", "word/settings.xml"]:
            xml_path = tmp_path / rel
            if xml_path.exists():
                xml = xml_path.read_text(encoding="utf-8")
                xml_path.write_text(replace_fonts_in_xml(xml), encoding="utf-8")

        output_docx.parent.mkdir(parents=True, exist_ok=True)
        if output_docx.exists():
            output_docx.unlink()
        zip_base = output_docx.with_suffix("")
        shutil.make_archive(str(zip_base), "zip", tmp_path)
        zip_base.with_suffix(".zip").rename(output_docx)


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a macOS LibreOffice PDF-safe DOCX with stable Chinese font names.")
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-docx", required=True)
    args = parser.parse_args()

    make_pdf_safe_docx(Path(args.input_docx), Path(args.output_docx))
    print(f"OK: {args.output_docx}")


if __name__ == "__main__":
    main()
