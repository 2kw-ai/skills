#!/usr/bin/env python3
"""Inspect, convert and read Word documents.

Three subcommands, each printing one JSON object on stdout:

  inspect FILE.docx   styles in use, headings, tables, sections, tracked changes and comments
  convert IN OUT      convert through LibreOffice; OUT's suffix picks the format (.docx, .pdf, .odt)
  text    FILE        the document as Markdown, through pandoc (.docx or .odt)

A failure prints {"error": "..."} and exits 1.
"""

import argparse
import json
import os
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

TARGETS = (".docx", ".pdf", ".odt")
PANDOC_READERS = {".docx": "docx", ".odt": "odt"}
LIST_LIMIT = 200  # per list in an inspect result; the counts stay exact
HEADING = re.compile(r"^(Heading (\d)|Title)$")


class ToolError(Exception):
    """A failure the model can act on; its message goes to stdout as {"error": ...}."""


def convert(source: str, target: str, timeout: int) -> dict:
    src, destination = Path(source), Path(target)
    if not src.is_file():
        raise ToolError(f"{source} does not exist")
    suffix = destination.suffix.lower()
    if suffix not in TARGETS:
        raise ToolError(f"cannot convert to {suffix or 'a file without a suffix'}; use one of {', '.join(TARGETS)}")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        raise ToolError("LibreOffice is not available here")
    # A fresh profile per call: two calls never share a profile lock, and nothing of the user's is read.
    work = Path(tempfile.mkdtemp(prefix="docx-convert-"))
    try:
        out_dir = work / "out"
        out_dir.mkdir()
        command = [
            soffice,
            f"-env:UserInstallation={(work / 'profile').as_uri()}",
            "--headless",
            "--convert-to",
            suffix[1:],
            "--outdir",
            str(out_dir),
            str(src.resolve()),
        ]
        produced = out_dir / (src.stem + suffix)
        deadline = time.monotonic() + timeout
        log = ""
        # A first start on a fresh profile can end without converting anything, so one retry is allowed.
        for _round in range(2):
            process = subprocess.Popen(
                command,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                start_new_session=True,
            )
            try:
                log, _ = process.communicate(timeout=max(deadline - time.monotonic(), 1))
            except subprocess.TimeoutExpired:
                if hasattr(os, "killpg"):
                    os.killpg(process.pid, signal.SIGKILL)  # soffice starts a child of its own
                else:
                    process.kill()
                process.communicate()
                raise ToolError(f"LibreOffice did not finish within {timeout} s; try a smaller document or --timeout")
            if process.returncode == 0 and produced.is_file():
                break
        if not produced.is_file():
            raise ToolError(f"LibreOffice could not convert {source}: {log.strip()[-300:]}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), destination)
    finally:
        shutil.rmtree(work, ignore_errors=True)
    return {"ok": True, "output": str(destination)}


def text(source: str, changes: str, output: str | None, timeout: int) -> dict:
    src = Path(source)
    if not src.is_file():
        raise ToolError(f"{source} does not exist")
    reader = PANDOC_READERS.get(src.suffix.lower())
    if reader is None:
        raise ToolError(f"pandoc does not read {src.suffix or 'this file'}; convert it to .docx first")
    if shutil.which("pandoc") is None:
        raise ToolError("pandoc is not available here")
    # --standalone keeps the front matter: pandoc moves Title, Subtitle and Author paragraphs out of the body.
    command = ["pandoc", "--from", reader, "--to", "markdown", "--wrap=none", "--standalone"]
    command += [f"--track-changes={changes}", str(src)]
    try:
        done = subprocess.run(
            command, check=False, capture_output=True, text=True, timeout=timeout, stdin=subprocess.DEVNULL
        )
    except subprocess.TimeoutExpired:
        raise ToolError(f"pandoc did not finish within {timeout} s")
    if done.returncode != 0:
        raise ToolError(f"pandoc could not read {source}: {done.stderr.strip()[-300:]}")
    markdown = done.stdout
    if output is None:
        return {"markdown": markdown, "characters": len(markdown)}
    destination = Path(output)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(markdown, encoding="utf-8")
    return {"ok": True, "output": str(destination), "characters": len(markdown)}


def inspect(path: str) -> dict:
    import docx
    from docx.oxml.ns import qn

    if not Path(path).is_file():
        raise ToolError(f"{path} does not exist")
    try:
        document = docx.Document(path)
    except Exception as error:  # noqa: BLE001 - a wrong format or a damaged file: say so, do not dump a traceback
        raise ToolError(
            f"{path} is not a .docx python-docx can read ({type(error).__name__}); a .doc or .odt needs convert first"
        )

    paragraph_styles: Counter = Counter()
    character_styles: Counter = Counter()

    def count(paragraphs) -> None:
        for paragraph in paragraphs:
            paragraph_styles[paragraph.style.name if paragraph.style is not None else None] += 1
            for run in paragraph.runs:
                if run.style is not None and run.style.name != "Default Paragraph Font":
                    character_styles[run.style.name] += 1

    headings = []
    for paragraph in document.paragraphs:
        name = paragraph.style.name if paragraph.style is not None else ""
        match = HEADING.match(name or "")
        if match and len(headings) < LIST_LIMIT:
            headings.append({"level": int(match.group(2) or 0), "style": name, "text": paragraph.text})
    count(document.paragraphs)

    tables = []
    for index, table in enumerate(document.tables):
        for row in table.rows:
            for cell in row.cells:
                count(cell.paragraphs)
        if len(tables) < LIST_LIMIT:
            first = [cell.text for cell in table.rows[0].cells] if table.rows else []
            tables.append(
                {
                    "index": index,
                    "rows": len(table.rows),
                    "columns": len(table.columns),
                    "style": table.style.name if table.style is not None else None,
                    "first_row": first,
                }
            )

    def mm(length) -> float | None:
        return None if length is None else round(length.mm, 1)

    def own_text(part) -> str | None:
        """A header's or footer's text; None when it repeats the previous section's."""
        if part.is_linked_to_previous:
            return None
        return "\n".join(paragraph.text for paragraph in part.paragraphs if paragraph.text)

    sections = [
        {
            "orientation": section.orientation.name.lower(),
            "page_width_mm": mm(section.page_width),
            "page_height_mm": mm(section.page_height),
            "margins_mm": {
                "top": mm(section.top_margin),
                "bottom": mm(section.bottom_margin),
                "left": mm(section.left_margin),
                "right": mm(section.right_margin),
            },
            "header_text": own_text(section.header),
            "footer_text": own_text(section.footer),
        }
        for section in document.sections
    ]

    # shortcut: tracked changes are counted in the body only, not in headers, footers or footnotes.
    body = document.element.body
    kinds = {"insertions": "w:ins", "deletions": "w:del", "moves": "w:moveTo", "format_changes": "w:rPrChange"}
    changes = {kind: sum(1 for _ in body.iter(qn(tag))) for kind, tag in kinds.items()}
    changes["format_changes"] += sum(1 for _ in body.iter(qn("w:pPrChange")))
    authors = sorted(
        {
            element.get(qn("w:author"))
            for tag in ("w:ins", "w:del", "w:moveTo", "w:moveFrom", "w:rPrChange", "w:pPrChange")
            for element in body.iter(qn(tag))
            if element.get(qn("w:author"))
        }
    )
    comments = [
        {"author": comment.author, "text": comment.text}
        for comment in list(document.comments)[:LIST_LIMIT]
    ]

    return {
        "paragraphs": len(document.paragraphs),
        "paragraph_styles": dict(paragraph_styles.most_common()),
        "character_styles": dict(character_styles.most_common()),
        "headings": headings,
        "tables": tables,
        "table_count": len(document.tables),
        "sections": sections,
        "tracked_changes": {**changes, "total": sum(changes.values()), "authors": authors},
        "comment_count": len(document.comments),
        "comments": comments,
    }


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser(
        "inspect", help="styles in use, headings, tables, sections, tracked changes, comments"
    ).add_argument("file")
    one = commands.add_parser(
        "convert", help="convert through LibreOffice: .doc/.odt/.rtf/.docx in; .docx, .pdf or .odt out (from OUT's suffix)"
    )
    one.add_argument("input")
    one.add_argument("output")
    one.add_argument("--timeout", type=int, default=120, help="seconds before LibreOffice is stopped (default 120)")
    two = commands.add_parser("text", help="the document as Markdown through pandoc (.docx or .odt)")
    two.add_argument("file")
    two.add_argument(
        "--changes",
        choices=("accept", "reject", "all"),
        default="accept",
        help="tracked changes: accept (default), reject, or all to show insertions, deletions and comments marked",
    )
    two.add_argument("--output", help="write the Markdown to this file instead of into the JSON")
    two.add_argument("--timeout", type=int, default=120, help="seconds before pandoc is stopped (default 120)")
    args = parser.parse_args(argv)
    try:
        if args.command == "convert":
            result = convert(args.input, args.output, args.timeout)
        elif args.command == "text":
            result = text(args.file, args.changes, args.output, args.timeout)
        else:
            result = inspect(args.file)
    except ToolError as error:
        print(json.dumps({"error": str(error)}, ensure_ascii=False))
        return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
