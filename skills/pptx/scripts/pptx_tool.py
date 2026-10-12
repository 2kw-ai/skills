#!/usr/bin/env python3
"""Inspect and render PowerPoint decks.

Two subcommands, each printing one JSON object on stdout:

  inspect FILE.pptx    slide size, layouts with their placeholders, and per slide: layout, title,
                       placeholders (empty ones flagged), all text, pictures, tables, charts, notes
  render  FILE OUTDIR  convert the deck to PDF and one PNG per slide (OUTDIR/slide-<n>.png)

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
from pathlib import Path

EMU_PER_INCH = 914400
SLIDE_PNG = re.compile(r"slide-(\d+)\.png")


class ToolError(Exception):
    """A failure the model can act on; its message goes to stdout as {"error": ...}."""


def load(path: str):
    from pptx import Presentation

    if not Path(path).is_file():
        raise ToolError(f"{path} does not exist")
    try:
        return Presentation(path)
    except Exception as error:  # noqa: BLE001 - a wrong format or a damaged file: say so, do not dump a traceback
        raise ToolError(
            f"{path} is not a .pptx deck python-pptx can read ({type(error).__name__}); "
            "a .ppt or .odp can still be rendered"
        )


def placeholder_type(shape) -> str:
    kind = shape.placeholder_format.type
    return kind.name if kind is not None else "UNKNOWN"


def flat(shapes):
    """The slide's shapes with groups opened up, in order."""
    from pptx.shapes.group import GroupShape

    for shape in shapes:
        if isinstance(shape, GroupShape):
            yield from flat(shape.shapes)
        else:
            yield shape


def shape_texts(shapes) -> list[str]:
    """Every piece of text on a slide in shape order: text frames, and tables row by row."""
    texts: list[str] = []
    for shape in flat(shapes):
        if getattr(shape, "has_table", False):
            texts += [" | ".join(cell.text for cell in row.cells) for row in shape.table.rows]
        elif shape.has_text_frame and shape.text_frame.text.strip():
            texts.append(shape.text_frame.text)
    return texts


def notes(slide) -> str:
    frame = slide.notes_slide.notes_text_frame if slide.has_notes_slide else None
    return frame.text if frame is not None else ""


def inspect(path: str) -> dict:
    deck = load(path)
    layouts = []
    for master_index, master in enumerate(deck.slide_masters):
        for layout in master.slide_layouts:
            layouts.append(
                {
                    "master": master_index,
                    "name": layout.name,
                    "placeholders": [
                        {"idx": ph.placeholder_format.idx, "type": placeholder_type(ph), "name": ph.name}
                        for ph in layout.placeholders
                    ],
                }
            )
    from pptx.shapes.picture import Picture

    slides = []
    for number, slide in enumerate(deck.slides, start=1):
        placeholders = []
        for ph in slide.placeholders:
            entry = {"idx": ph.placeholder_format.idx, "type": placeholder_type(ph), "name": ph.name}
            # A filled picture, chart or table placeholder has no text frame; an unfilled one of any kind has
            # an empty one and shows "Click to add ..." while editing: fill it or delete it.
            entry["empty"] = ph.has_text_frame and not ph.text_frame.text.strip()
            if ph.has_text_frame and not entry["empty"]:
                entry["text"] = ph.text_frame.text
            placeholders.append(entry)
        shapes = list(flat(slide.shapes))
        title = slide.shapes.title
        slides.append(
            {
                "number": number,
                "layout": slide.slide_layout.name,
                "title": title.text_frame.text if title is not None and title.has_text_frame else None,
                "placeholders": placeholders,
                "text": shape_texts(slide.shapes),
                "pictures": sum(isinstance(shape, Picture) for shape in shapes),
                "tables": sum(bool(getattr(shape, "has_table", False)) for shape in shapes),
                "charts": sum(bool(getattr(shape, "has_chart", False)) for shape in shapes),
                "notes": notes(slide),
            }
        )
    return {
        "slide_width_in": round(deck.slide_width / EMU_PER_INCH, 2),
        "slide_height_in": round(deck.slide_height / EMU_PER_INCH, 2),
        "layouts": layouts,
        "slide_count": len(slides),
        "slides": slides,
    }


def to_pdf(src: Path, work: Path, deadline: float, timeout: int) -> Path:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice is None:
        raise ToolError("LibreOffice is not available here")
    # A fresh profile per call: two calls never share a profile lock, and nothing of the user's is read.
    (work / "profile").mkdir()
    command = [
        soffice,
        f"-env:UserInstallation={(work / 'profile').as_uri()}",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(work),
        str(src.resolve()),
    ]
    produced = work / (src.stem + ".pdf")
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
            raise ToolError(f"LibreOffice did not finish within {timeout} s; try a smaller deck or --timeout")
        if process.returncode == 0 and produced.is_file():
            return produced
    raise ToolError(f"LibreOffice could not convert {src}: {log.strip()[-300:]}")


def slide_pngs(folder: Path) -> list[Path]:
    """slide-<n>.png in slide order; pdftoppm pads n to the width of the page count (slide-01.png from 10 pages)."""
    found = [(int(m.group(1)), p) for p in folder.iterdir() if (m := SLIDE_PNG.fullmatch(p.name))]
    return [p for _n, p in sorted(found)]


def render(source: str, out: str, dpi: int, timeout: int) -> dict:
    src = Path(source)
    if not src.is_file():
        raise ToolError(f"{source} does not exist")
    if shutil.which("pdftoppm") is None:
        raise ToolError("pdftoppm is not available here")
    deadline = time.monotonic() + timeout
    out_dir = Path(out)
    out_dir.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix="pptx-render-"))
    try:
        pdf = to_pdf(src, work, deadline, timeout)
        for old in slide_pngs(out_dir):  # a slide deleted since the last render must not linger
            old.unlink()
        try:
            done = subprocess.run(
                ["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(out_dir / "slide")],
                check=False,
                capture_output=True,
                text=True,
                stdin=subprocess.DEVNULL,
                timeout=max(deadline - time.monotonic(), 1),
            )
        except subprocess.TimeoutExpired:
            raise ToolError(f"rendering did not finish within {timeout} s; try a lower --dpi or --timeout")
        if done.returncode != 0:
            raise ToolError(f"pdftoppm could not render the slides: {done.stderr.strip()[-300:]}")
        target_pdf = out_dir / (src.stem + ".pdf")
        shutil.move(str(pdf), target_pdf)
    finally:
        shutil.rmtree(work, ignore_errors=True)
    images = slide_pngs(out_dir)
    return {"ok": True, "pdf": str(target_pdf), "slides": len(images), "images": [str(p) for p in images]}


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser(
        "inspect", help="slide size, layouts, placeholders and text per slide (.pptx only)"
    ).add_argument("file")
    one = commands.add_parser(
        "render", help="deck (.pptx, .ppt, .odp) to OUTDIR/<name>.pdf and OUTDIR/slide-<n>.png"
    )
    one.add_argument("input")
    one.add_argument("outdir")
    one.add_argument("--dpi", type=int, default=100, help="PNG resolution (default 100: 1000 px across a 10 in slide)")
    one.add_argument("--timeout", type=int, default=180, help="seconds before rendering is stopped (default 180)")
    args = parser.parse_args(argv)
    try:
        if args.command == "inspect":
            result = inspect(args.file)
        else:
            result = render(args.input, args.outdir, args.dpi, args.timeout)
    except ToolError as error:
        print(json.dumps({"error": str(error)}, ensure_ascii=False))
        return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
