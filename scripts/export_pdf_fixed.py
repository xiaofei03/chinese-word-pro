from __future__ import annotations

import argparse
import os
import shutil
import subprocess
from pathlib import Path


DEFAULT_SOFFICE = Path("/Users/xiaofei/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/soffice")
DEFAULT_PDFTOPPM = Path("/Users/xiaofei/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/pdftoppm")
SKILL_DIR = Path(__file__).resolve().parents[1]
MAKE_SAFE = SKILL_DIR / "scripts" / "make_pdf_safe_docx.py"


def write_fontconfig(fontconfig_file: Path) -> None:
    fontconfig_file.parent.mkdir(parents=True, exist_ok=True)
    fontconfig_file.write_text(
        """<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "fonts.dtd">
<fontconfig>
  <dir>/System/Library/Fonts</dir>
  <dir>/System/Library/Fonts/Supplemental</dir>
  <dir>/Library/Fonts</dir>
  <dir>~/Library/Fonts</dir>
  <alias>
    <family>宋体</family>
    <prefer><family>宋体-简</family></prefer>
  </alias>
  <alias>
    <family>黑体</family>
    <prefer><family>黑体-简</family></prefer>
  </alias>
  <alias>
    <family>Times New Roman</family>
    <prefer><family>Times New Roman</family></prefer>
  </alias>
</fontconfig>
""",
        encoding="utf-8",
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Export a Chinese DOCX to PDF with stable macOS fonts and render PNG pages for QA.")
    parser.add_argument("--input-docx", required=True)
    parser.add_argument("--output-pdf", required=True)
    parser.add_argument("--work-dir", default=None, help="Directory for intermediate PDF-safe DOCX and fontconfig.")
    parser.add_argument("--render-dir", default=None, help="Optional directory for rendered PNG pages.")
    parser.add_argument("--soffice", default=str(DEFAULT_SOFFICE))
    parser.add_argument("--pdftoppm", default=str(DEFAULT_PDFTOPPM))
    args = parser.parse_args()

    input_docx = Path(args.input_docx).resolve()
    output_pdf = Path(args.output_pdf).resolve()
    work_dir = Path(args.work_dir).resolve() if args.work_dir else output_pdf.parent / "_pdf_export_work"
    render_dir = Path(args.render_dir).resolve() if args.render_dir else None
    safe_docx = work_dir / f"{input_docx.stem}_PDF安全版.docx"
    fontconfig_file = work_dir / "fontconfig" / "fonts.conf"

    work_dir.mkdir(parents=True, exist_ok=True)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    write_fontconfig(fontconfig_file)

    subprocess.run(
        [os.environ.get("PYTHON", "python3"), str(MAKE_SAFE), "--input-docx", str(input_docx), "--output-docx", str(safe_docx)],
        check=True,
    )

    export_dir = work_dir / "pdf"
    shutil.rmtree(export_dir, ignore_errors=True)
    export_dir.mkdir(parents=True, exist_ok=True)

    env = os.environ.copy()
    env["FONTCONFIG_FILE"] = str(fontconfig_file)
    subprocess.run(
        [args.soffice, "--headless", "--convert-to", "pdf", "--outdir", str(export_dir), str(safe_docx)],
        check=True,
        env=env,
    )

    produced = export_dir / f"{safe_docx.stem}.pdf"
    if output_pdf.exists():
        output_pdf.unlink()
    produced.rename(output_pdf)

    if render_dir:
        shutil.rmtree(render_dir, ignore_errors=True)
        render_dir.mkdir(parents=True, exist_ok=True)
        subprocess.run([args.pdftoppm, "-png", "-r", "150", str(output_pdf), str(render_dir / "page")], check=True)

    print(f"OK: {output_pdf}")
    if render_dir:
        print(f"OK: {render_dir}")


if __name__ == "__main__":
    main()
