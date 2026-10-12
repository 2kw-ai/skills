#!/usr/bin/env python3
"""Look into PDF files, pull their text and tables out, fill their forms and render their pages.

Five subcommands, each printing one JSON object on stdout:

  info    FILE                 pages, text layer per page, form fields, encryption
  extract FILE                 text and tables per page; --ocr reads pages that have no text layer,
                               --ocr-pages forces OCR on chosen pages
  fields  FILE                 the fillable form fields: name, type, value, allowed values, page
  fill    FILE VALUES.json OUT fill form fields from a JSON object {field name: value}; --flatten on request
  render  FILE OUTDIR          one PNG per page, to look at

A failure prints {"error": "..."} and exits 1.
"""

import argparse
import json
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import time
from pathlib import Path

FIELD_TYPES = {"/Tx": "text", "/Btn": "button", "/Ch": "choice", "/Sig": "signature"}
RADIO_FLAG = 1 << 15
PUSHBUTTON_FLAG = 1 << 16
OCR_DPI = 300


class ToolError(Exception):
    """A failure the model can act on; its message goes to stdout as {"error": ...}."""


def run(command: list[str], deadline: float, timeout: int) -> str:
    """Run one external program inside the call's time budget; a hang is killed with everything it started."""
    process = subprocess.Popen(
        command,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        start_new_session=True,
    )
    try:
        out, err = process.communicate(timeout=max(deadline - time.monotonic(), 0.1))
    except subprocess.TimeoutExpired:
        if hasattr(os, "killpg"):
            os.killpg(process.pid, signal.SIGKILL)
        else:
            process.kill()
        process.communicate()
        raise ToolError(f"{Path(command[0]).name} did not finish within {timeout} s; try fewer pages or --timeout")
    if process.returncode != 0:
        raise ToolError(f"{Path(command[0]).name} failed: {err.strip()[-300:]}")
    return out


def need(program: str) -> str:
    found = shutil.which(program)
    if found is None:
        raise ToolError(f"{program} is not available here")
    return found


def open_reader(path: str, password: str | None):
    from pypdf import PdfReader

    if not Path(path).is_file():
        raise ToolError(f"{path} does not exist")
    try:
        reader = PdfReader(path)
    except Exception as error:  # noqa: BLE001 - a wrong format or a damaged file: say so, do not dump a traceback
        raise ToolError(f"{path} is not a PDF pypdf can read ({type(error).__name__}: {error})")
    if reader.is_encrypted:
        try:
            unlocked = reader.decrypt(password or "")
        except Exception as error:  # noqa: BLE001 - e.g. an encryption method without its crypto library
            raise ToolError(f"{path} is encrypted and could not be opened ({error})")
        if not unlocked:
            raise ToolError(f"{path} is encrypted; pass --password (a wrong or missing password)")
    return reader


def page_numbers(spec: str | None, count: int) -> list[int]:
    """'1-3,5' -> [1, 2, 3, 5], 1-based; None means every page."""
    if not spec:
        return list(range(1, count + 1))
    numbers: list[int] = []
    for part in spec.split(","):
        first, _, last = part.strip().partition("-")
        try:
            start, end = int(first), int(last or first)
        except ValueError:
            raise ToolError(f"--pages {spec!r}: write page numbers like 1-3,5")
        if not 1 <= start <= end <= count:
            raise ToolError(f"--pages {spec!r}: the file has {count} page(s)")
        numbers.extend(n for n in range(start, end + 1) if n not in numbers)
    return numbers


def widget_name(widget) -> str:
    """The fully qualified field name of a widget annotation: its /T and its parents' /T joined with dots."""
    parts = []
    node = widget
    while node is not None:
        if "/T" in node:
            parts.append(str(node["/T"]))
        node = node["/Parent"].get_object() if "/Parent" in node else None
    return ".".join(reversed(parts))


def field_pages(reader) -> dict[str, list[int]]:
    pages: dict[str, list[int]] = {}
    for number, page in enumerate(reader.pages, start=1):
        for annotation in page.get("/Annots") or []:
            widget = annotation.get_object()
            if widget.get("/Subtype") == "/Widget":
                listed = pages.setdefault(widget_name(widget), [])
                if number not in listed:
                    listed.append(number)
    return pages


def describe_fields(reader) -> list[dict]:
    fields = reader.get_fields() or {}
    pages = field_pages(reader)
    described = []
    for name, field in fields.items():
        kind = FIELD_TYPES.get(field.get("/FT"), "unknown")
        if kind == "unknown" and name not in pages:
            continue  # a pure container of other fields
        flags = int(field.get("/Ff", 0))
        if kind == "button":
            kind = "radio" if flags & RADIO_FLAG else "pushbutton" if flags & PUSHBUTTON_FLAG else "checkbox"
        value = field.get("/V")
        entry = {"name": name, "type": kind, "value": None if value is None else str(value)}
        if isinstance(value, list):
            entry["value"] = [str(item) for item in value]
        if kind in ("checkbox", "radio"):
            entry["options"] = [str(state) for state in field.get("/_States_", []) if str(state) != "/Off"]
        elif kind == "choice":
            entry["options"] = [str(o[-1]) if isinstance(o, list) else str(o) for o in field.get("/Opt", [])]
        if field.get("/TU"):
            entry["tooltip"] = str(field["/TU"])
        entry["pages"] = pages.get(name, [])
        described.append(entry)
    return described


def info(path: str, password: str | None) -> dict:
    from pypdf import PdfReader

    if not Path(path).is_file():
        raise ToolError(f"{path} does not exist")
    try:
        encrypted = PdfReader(path).is_encrypted
    except Exception as error:  # noqa: BLE001
        raise ToolError(f"{path} is not a PDF pypdf can read ({type(error).__name__}: {error})")
    try:
        reader = open_reader(path, password)
    except ToolError as error:
        if not encrypted:
            raise
        return {"encrypted": True, "locked": True, "hint": str(error)}
    pages = []
    for number, page in enumerate(reader.pages, start=1):
        text = (page.extract_text() or "").strip()
        box = page.mediabox
        pages.append(
            {
                "page": number,
                "width_pt": round(float(box.width), 1),
                "height_pt": round(float(box.height), 1),
                "rotation": page.rotation,
                "text_chars": len(text),
                "has_text": bool(text),
            }
        )
    metadata = reader.metadata or {}
    return {
        "pages": len(pages),
        "encrypted": encrypted,
        "title": metadata.get("/Title"),
        "author": metadata.get("/Author"),
        "form_fields": len(describe_fields(reader)),
        "pages_without_text": [p["page"] for p in pages if not p["has_text"]],
        "page_details": pages,
    }


def ocr_page(path: str, number: int, lang: str, password: str | None, work: Path, deadline: float, timeout: int) -> str:
    image = work / f"ocr-{number}"
    command = [need("pdftoppm"), "-r", str(OCR_DPI), "-gray", "-png", "-singlefile", "-f", str(number), "-l", str(number)]
    if password:
        command += ["-upw", password]
    run(command + [path, str(image)], deadline, timeout)
    return run([need("tesseract"), f"{image}.png", "stdout", "-l", lang], deadline, timeout)


def extract(
    path: str, pages_spec: str | None, ocr: bool, ocr_spec: str | None, lang: str, password: str | None, timeout: int
) -> dict:
    import pdfplumber

    open_reader(path, password)  # the same clear errors for a missing, damaged or locked file
    deadline = time.monotonic() + timeout
    result = []
    work = Path(tempfile.mkdtemp(prefix="pdf-extract-"))
    try:
        with pdfplumber.open(path, password=password or "") as document:
            forced = set(page_numbers(ocr_spec, len(document.pages))) if ocr_spec else set()
            for number in page_numbers(pages_spec, len(document.pages)):
                page = document.pages[number - 1]
                text = (page.extract_text() or "").strip()
                tables = page.extract_tables() if text else []
                source = "text" if text else "none"
                # A scan with a stamped page number has a little text, so only a forced page gets OCR then.
                if number in forced or (not text and ocr):
                    text = ocr_page(path, number, lang, password, work, deadline, timeout).strip()
                    source, tables = "ocr", []
                elif time.monotonic() > deadline:
                    raise ToolError(f"extract did not finish within {timeout} s; try fewer pages or --timeout")
                result.append({"page": number, "source": source, "text": text, "tables": tables})
    finally:
        shutil.rmtree(work, ignore_errors=True)
    answer: dict = {"pages": result}
    missing = [p["page"] for p in result if p["source"] == "none"]
    if missing:
        answer["pages_without_text"] = missing
        answer["hint"] = "these pages have no text layer (a scan?); run again with --ocr"
    return answer


def fields(path: str, password: str | None) -> dict:
    return {"fields": describe_fields(open_reader(path, password))}


def fill(path: str, values_path: str, target: str, flatten: bool, password: str | None) -> dict:
    from pypdf import PdfWriter
    from pypdf.generic import NameObject

    reader = open_reader(path, password)
    try:
        values = json.loads(Path(values_path).read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise ToolError(f"{values_path} does not exist")
    except json.JSONDecodeError as error:
        raise ToolError(f"{values_path} is not valid JSON: {error}")
    if not isinstance(values, dict) or not values:
        raise ToolError(f'{values_path} must hold a JSON object {{"field name": value, ...}}')
    known = {field["name"]: field for field in describe_fields(reader)}
    if not known:
        raise ToolError(f"{path} has no fillable form fields; write the text onto the page instead")
    unknown = sorted(set(values) - set(known))
    if unknown:
        raise ToolError(f"unknown field(s) {unknown}; the file has {sorted(known)[:50]}")
    prepared: dict[str, str] = {}
    problems = []
    for name, value in values.items():
        field = known[name]
        if field["type"] in ("checkbox", "radio"):
            options = field["options"]
            if value is True and field["type"] == "checkbox" and options:
                value = options[0]
            elif value in (False, None):
                value = "/Off"
            elif isinstance(value, str) and "/" + value.lstrip("/") in options:
                value = "/" + value.lstrip("/")
            else:
                problems.append(f"{name}: {value!r} is not one of {options} (or true/false for a checkbox)")
                continue
        elif field["type"] in ("pushbutton", "signature"):
            problems.append(f"{name}: a {field['type']} field cannot be filled with a value")
            continue
        elif field["type"] == "choice" and field["options"] and str(value) not in field["options"]:
            problems.append(f"{name}: {value!r} is not one of {field['options']}")
            continue
        prepared[name] = "" if value is None else str(value)
    if problems:
        raise ToolError("; ".join(problems))
    writer = PdfWriter(clone_from=reader)
    writer.update_page_form_field_values(None, prepared, auto_regenerate=not flatten, flatten=flatten)
    if flatten:
        # The values are now drawn into the page; the widgets and the form itself go, so nothing stays editable.
        writer.remove_annotations("/Widget")
        if "/AcroForm" in writer._root_object:
            del writer._root_object[NameObject("/AcroForm")]
    destination = Path(target)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with open(destination, "wb") as handle:
        writer.write(handle)
    return {"ok": True, "output": str(destination), "filled": sorted(prepared), "flattened": flatten}


def render(path: str, out_dir: str, pages_spec: str | None, dpi: int, password: str | None, timeout: int) -> dict:
    reader = open_reader(path, password)
    deadline = time.monotonic() + timeout
    target = Path(out_dir)
    target.mkdir(parents=True, exist_ok=True)
    stem = Path(path).stem
    images = []
    for number in page_numbers(pages_spec, len(reader.pages)):
        image = target / f"{stem}-page-{number}"
        command = [need("pdftoppm"), "-r", str(dpi), "-png", "-singlefile", "-f", str(number), "-l", str(number)]
        if password:
            command += ["-upw", password]
        run(command + [path, str(image)], deadline, timeout)
        images.append({"page": number, "png": f"{image}.png"})
    return {"ok": True, "images": images}


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)

    def command(name: str, help_text: str) -> argparse.ArgumentParser:
        sub = commands.add_parser(name, help=help_text, description=help_text)
        sub.add_argument("file")
        sub.add_argument("--password", help="password of an encrypted file")
        return sub

    command("info", "pages, text layer per page, form fields, encryption")
    one = command("extract", "text and tables per page as JSON")
    one.add_argument("--pages", help="pages to read, like 1-3,5 (default: all)")
    one.add_argument("--ocr", action="store_true", help="OCR pages that have no text layer")
    one.add_argument("--ocr-pages", help="OCR these pages even when they have some text, like 2,4-5")
    one.add_argument("--lang", default="deu+eng", help="OCR languages (default deu+eng)")
    one.add_argument("--timeout", type=int, default=300, help="seconds for the whole call (default 300)")
    command("fields", "list the fillable form fields")
    one = command("fill", "fill form fields from a JSON object {field name: value}; write OUTPUT")
    one.add_argument("values", help="JSON file; checkboxes take true/false, radios one of their options")
    one.add_argument("output")
    one.add_argument("--flatten", action="store_true", help="draw the values into the page; the form is gone")
    one = command("render", "one PNG per page into OUTDIR")
    one.add_argument("outdir")
    one.add_argument("--pages", help="pages to render, like 1-3,5 (default: all)")
    one.add_argument("--dpi", type=int, default=100, help="resolution (default 100; 150-200 for small print)")
    one.add_argument("--timeout", type=int, default=120, help="seconds for the whole call (default 120)")
    args = parser.parse_args(argv)
    try:
        if args.command == "info":
            result = info(args.file, args.password)
        elif args.command == "extract":
            result = extract(args.file, args.pages, args.ocr, args.ocr_pages, args.lang, args.password, args.timeout)
        elif args.command == "fields":
            result = fields(args.file, args.password)
        elif args.command == "fill":
            result = fill(args.file, args.values, args.output, args.flatten, args.password)
        else:
            result = render(args.file, args.outdir, args.pages, args.dpi, args.password, args.timeout)
    except ToolError as error:
        print(json.dumps({"error": str(error)}, ensure_ascii=False))
        return 1
    print(json.dumps(result, ensure_ascii=False, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
