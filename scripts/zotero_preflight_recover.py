#!/usr/bin/env python3
"""Zotero preflight and recovery helper for citation-safe Word delivery.

The script deliberately does not create a manuscript export. It only verifies
that Zotero is reachable enough for a downstream citation-aware export, and it
can open Zotero plus a target collection through Zotero's URI scheme.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
import urllib.error
import urllib.request
import zipfile
from pathlib import Path


def run(cmd: list[str], *, check: bool = False) -> subprocess.CompletedProcess[str]:
    return subprocess.run(cmd, text=True, capture_output=True, check=check)


def locate_zotero() -> Path | None:
    candidates = [
        Path("/Applications/Zotero.app"),
        Path.home() / "Applications/Zotero.app",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate

    md = run(
        [
            "mdfind",
            "kMDItemCFBundleIdentifier == 'org.zotero.zotero'",
        ]
    )
    for line in md.stdout.splitlines():
        candidate = Path(line.strip())
        if candidate.exists() and candidate.name == "Zotero.app":
            return candidate
    return None


def zotero_is_running() -> bool:
    return run(["pgrep", "-x", "Zotero"]).returncode == 0


def open_zotero(app_path: Path | None) -> None:
    if app_path:
        run(["open", str(app_path)], check=False)
    else:
        run(["open", "-a", "Zotero"], check=False)


def connector_ping(timeout: float = 2.0) -> tuple[bool, str]:
    urls = [
        "http://127.0.0.1:23119/connector/ping",
        "http://localhost:23119/connector/ping",
    ]
    for url in urls:
        for method in ("GET", "POST"):
            try:
                req = urllib.request.Request(url, method=method)
                with urllib.request.urlopen(req, timeout=timeout) as resp:
                    body = resp.read(200).decode("utf-8", errors="replace")
                    if resp.status < 500:
                        return True, f"{method} {url} -> {resp.status} {body[:80]}"
            except urllib.error.HTTPError as exc:
                if exc.code < 500:
                    return True, f"{method} {url} -> HTTP {exc.code}"
            except Exception as exc:  # noqa: BLE001 - report the probe failure
                last = f"{method} {url} -> {type(exc).__name__}: {exc}"
    return False, last if "last" in locals() else "no connector probe attempted"


def wait_for_zotero(timeout: int) -> tuple[bool, list[str]]:
    deadline = time.time() + timeout
    log: list[str] = []
    while time.time() < deadline:
        running = zotero_is_running()
        ok, msg = connector_ping()
        log.append(f"running={running}; connector={ok}; {msg}")
        # The connector is the capability needed by export workflows. On macOS
        # the process name can vary, so do not fail when pgrep misses Zotero.
        if ok:
            return True, log
        time.sleep(2)
    return False, log


def collection_uri(args: argparse.Namespace) -> str | None:
    if args.select_uri:
        return args.select_uri
    if not args.collection_key:
        return None
    if args.group_id:
        return f"zotero://select/groups/{args.group_id}/collections/{args.collection_key}"
    return f"zotero://select/library/collections/{args.collection_key}"


def inspect_smoke_docx(path: Path) -> tuple[bool, str]:
    if not path.exists():
        return False, f"smoke DOCX does not exist: {path}"
    try:
        with zipfile.ZipFile(path) as zf:
            xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
    except Exception as exc:  # noqa: BLE001
        return False, f"cannot inspect smoke DOCX: {type(exc).__name__}: {exc}"

    has_addin = "ADDIN ZOTERO_ITEM" in xml
    has_csl = "CSL_CITATION" in xml
    if has_addin and has_csl:
        return True, "smoke DOCX contains ADDIN ZOTERO_ITEM and CSL_CITATION"
    return False, (
        "smoke DOCX is missing live citation markers "
        f"(ADDIN_ZOTERO_ITEM={has_addin}, CSL_CITATION={has_csl})"
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Open Zotero, optionally select a collection, and verify Zotero readiness."
    )
    parser.add_argument("--collection-key", help="Zotero collection key to select.")
    parser.add_argument("--group-id", help="Zotero group id for group-library collections.")
    parser.add_argument(
        "--select-uri",
        help="Explicit zotero://select/... URI. Overrides --collection-key and --group-id.",
    )
    parser.add_argument(
        "--smoke-docx",
        type=Path,
        help="Optional tiny DOCX export to audit for live Zotero citation fields.",
    )
    parser.add_argument("--timeout", type=int, default=60, help="Seconds to wait for Zotero.")
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Do not launch Zotero; only check readiness.",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Fail if no collection URI was provided.",
    )
    args = parser.parse_args()

    app_path = locate_zotero()
    print(f"Zotero app: {app_path if app_path else 'not found by path search'}")

    if not args.no_open:
        open_zotero(app_path)

    ok, log = wait_for_zotero(args.timeout)
    for line in log[-5:]:
        print(f"probe: {line}")
    if not ok:
        print("ERROR: Zotero did not become reachable before timeout.", file=sys.stderr)
        return 2

    uri = collection_uri(args)
    if uri:
        print(f"Opening Zotero collection URI: {uri}")
        run(["open", uri], check=False)
        time.sleep(3)
    elif args.strict:
        print("ERROR: --strict requires --collection-key or --select-uri.", file=sys.stderr)
        return 2
    else:
        print("No collection URI provided; skipping collection selection.")

    ok, msg = connector_ping()
    print(f"final connector probe: {msg}")
    if not ok:
        print("ERROR: Zotero connector is not reachable after recovery.", file=sys.stderr)
        return 2

    if args.smoke_docx:
        smoke_ok, smoke_msg = inspect_smoke_docx(args.smoke_docx)
        print(smoke_msg)
        if not smoke_ok:
            return 2

    print("Zotero preflight passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
