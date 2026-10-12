"""The pdf skill's helper, end to end.

Needs Python 3 with pypdf, pdfplumber, reportlab and Pillow, plus pdftoppm, tesseract (deu, eng) and LibreOffice
(`soffice`) on PATH. SKILL_DIR names the bundle (default skills/pdf next to this repository's tests). The fixtures
are built at run time, so no binary file is committed.
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pdfplumber
from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet

TOOL = Path(os.environ.get("SKILL_DIR", Path(__file__).resolve().parents[2] / "skills" / "pdf")) / "scripts" / "pdf_tool.py"


def tool(*args: str) -> tuple[int, dict]:
    done = subprocess.run(
        [sys.executable, "-I", str(TOOL), *args],
        check=False,
        capture_output=True,
        text=True,
        timeout=240,
        stdin=subprocess.DEVNULL,
    )
    return done.returncode, json.loads(done.stdout)


class PdfToolTest(unittest.TestCase):
    def setUp(self) -> None:
        self.dir = Path(tempfile.mkdtemp(prefix="pdf-skill-"))

    def report(self, name: str = "report.pdf", pages: int = 2) -> Path:
        """Page 1: a heading and a ruled table; further pages: one line of text each."""
        path = self.dir / name
        styles = getSampleStyleSheet()
        table = Table([["Region", "Amount"], ["North", "120"], ["South", "80"]])
        table.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 1, colors.black)]))
        story = [Paragraph("Quarterly report", styles["Title"]), table]
        for number in range(2, pages + 1):
            story += [PageBreak(), Paragraph(f"Notes on page {number}", styles["Normal"])]
        SimpleDocTemplate(str(path), pagesize=A4).build(story)
        return path

    def scan(self, name: str = "scan.pdf", text: str = "Rechnung Nr 2026 Gesamtbetrag") -> Path:
        """A page that is only a picture of text, as a scanner makes it: no text layer at all."""
        image = Image.new("L", (1700, 2200), 255)
        try:
            font = ImageFont.truetype("DejaVuSans.ttf", 64)
        except OSError:
            font = ImageFont.load_default(size=64)
        ImageDraw.Draw(image).text((150, 300), text, fill=0, font=font)
        path = self.dir / name
        image.save(path, "PDF", resolution=200)
        return path

    def mixed(self) -> Path:
        """Page 1 has text, page 2 is a scan."""
        writer = PdfWriter()
        writer.append(str(self.report(pages=1)))
        writer.append(str(self.scan()))
        path = self.dir / "mixed.pdf"
        writer.write(path)
        return path

    def form(self, name: str = "form.pdf") -> Path:
        path = self.dir / name
        sheet = canvas.Canvas(str(path), pagesize=A4)
        sheet.drawString(50, 800, "Application")
        form = sheet.acroForm
        form.textfield(name="full_name", x=50, y=700, width=300, height=24, tooltip="Your name")
        form.checkbox(name="agree", x=50, y=650)
        form.choice(name="country", value="DE", options=["DE", "AT", "CH"], x=50, y=600, width=120, height=24)
        form.radio(name="size", value="small", selected=True, x=50, y=550)
        form.radio(name="size", value="large", x=100, y=550)
        sheet.save()
        return path

    def values(self, data: dict) -> Path:
        path = self.dir / "values.json"
        path.write_text(json.dumps(data), encoding="utf-8")
        return path

    # info

    def test_info_reports_pages_text_layer_and_fields(self) -> None:
        code, result = tool("info", str(self.mixed()))

        self.assertEqual(code, 0, result)
        self.assertEqual((result["pages"], result["encrypted"], result["form_fields"]), (2, False, 0))
        self.assertEqual(result["pages_without_text"], [2])
        self.assertTrue(result["page_details"][0]["has_text"])

        code, result = tool("info", str(self.form()))
        self.assertEqual((code, result["form_fields"]), (0, 4), result)

    def test_info_on_an_encrypted_file_says_so_and_opens_with_the_password(self) -> None:
        writer = PdfWriter(clone_from=str(self.report()))
        writer.encrypt("secret", algorithm="RC4-128")
        locked = self.dir / "locked.pdf"
        writer.write(locked)

        code, result = tool("info", str(locked))
        self.assertEqual((code, result["encrypted"], result["locked"]), (0, True, True), result)

        code, result = tool("info", str(locked), "--password", "secret")
        self.assertEqual((code, result["encrypted"], result["pages"]), (0, True, 2), result)

        code, result = tool("extract", str(locked))
        self.assertEqual(code, 1)
        self.assertIn("--password", result["error"])

    # extract

    def test_extract_returns_text_and_tables_per_page(self) -> None:
        code, result = tool("extract", str(self.report()))

        self.assertEqual(code, 0, result)
        first, second = result["pages"]
        self.assertIn("Quarterly report", first["text"])
        self.assertEqual(first["tables"], [[["Region", "Amount"], ["North", "120"], ["South", "80"]]])
        self.assertEqual((second["page"], second["source"]), (2, "text"))
        self.assertNotIn("pages_without_text", result)

        code, result = tool("extract", str(self.report()), "--pages", "2")
        self.assertEqual([p["page"] for p in result["pages"]], [2])

    def test_extract_points_at_ocr_for_a_scan_and_ocr_reads_it(self) -> None:
        path = self.mixed()

        code, result = tool("extract", str(path))
        self.assertEqual(code, 0, result)
        self.assertEqual((result["pages"][1]["source"], result["pages_without_text"]), ("none", [2]))
        self.assertIn("--ocr", result["hint"])

        code, result = tool("extract", str(path), "--ocr")
        self.assertEqual(code, 0, result)
        self.assertEqual(result["pages"][0]["source"], "text", "a page with text is not OCRed")
        scanned = result["pages"][1]
        self.assertEqual(scanned["source"], "ocr")
        for word in ("Rechnung", "2026", "Gesamtbetrag"):
            self.assertIn(word, scanned["text"])

    def test_ocr_pages_forces_ocr_on_a_scan_with_a_stamped_page_number(self) -> None:
        stamp = self.dir / "stamp.pdf"
        sheet = canvas.Canvas(str(stamp), pagesize=(612, 792))  # the scan's page size: 1700 x 2200 px at 200 dpi
        sheet.drawString(300, 20, "7")
        sheet.save()
        writer = PdfWriter(clone_from=str(self.scan()))
        writer.pages[0].merge_page(PdfReader(stamp).pages[0])
        path = self.dir / "stamped.pdf"
        writer.write(path)

        code, result = tool("extract", str(path), "--ocr")
        self.assertEqual(code, 0, result)
        self.assertEqual((result["pages"][0]["source"], result["pages"][0]["text"]), ("text", "7"))

        code, result = tool("extract", str(path), "--ocr-pages", "1")
        self.assertEqual(code, 0, result)
        self.assertEqual(result["pages"][0]["source"], "ocr")
        self.assertIn("Gesamtbetrag", result["pages"][0]["text"])

    def test_ocr_calls_in_parallel_do_not_share_files(self) -> None:
        scans = [self.scan(f"scan{i}.pdf", f"Lieferschein {i}0{i}") for i in (1, 2)]

        with ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(lambda path: tool("extract", str(path), "--ocr"), scans))

        for index, (code, result) in zip((1, 2), results):
            self.assertEqual(code, 0, result)
            self.assertIn(f"{index}0{index}", result["pages"][0]["text"])

    def test_extract_rejects_pages_outside_the_file(self) -> None:
        code, result = tool("extract", str(self.report()), "--pages", "3")
        self.assertEqual(code, 1)
        self.assertIn("2 page(s)", result["error"])

    # fields and fill

    def test_fields_lists_name_type_options_and_page(self) -> None:
        code, result = tool("fields", str(self.form()))

        self.assertEqual(code, 0, result)
        by_name = {field["name"]: field for field in result["fields"]}
        self.assertEqual(sorted(by_name), ["agree", "country", "full_name", "size"])
        self.assertEqual((by_name["full_name"]["type"], by_name["full_name"]["pages"]), ("text", [1]))
        self.assertEqual(by_name["full_name"]["tooltip"], "Your name")
        self.assertEqual(by_name["agree"]["type"], "checkbox")
        self.assertEqual(by_name["country"]["options"], ["DE", "AT", "CH"])
        self.assertEqual((by_name["size"]["type"], sorted(by_name["size"]["options"])), ("radio", ["/large", "/small"]))

    def test_fill_sets_values_that_read_back(self) -> None:
        target = self.dir / "out" / "filled.pdf"
        data = {"full_name": "Jürgen Weiß", "agree": True, "country": "AT", "size": "large"}

        code, result = tool("fill", str(self.form()), str(self.values(data)), str(target))

        self.assertEqual(code, 0, result)
        self.assertEqual(result["filled"], sorted(data))
        read = {name: field.get("/V") for name, field in PdfReader(target).get_fields().items()}
        self.assertEqual(read["full_name"], "Jürgen Weiß")
        self.assertEqual(read["country"], "AT")
        self.assertEqual(read["size"], "/large")
        self.assertNotEqual(read["agree"], "/Off")

    def test_fill_flatten_leaves_the_text_and_no_form(self) -> None:
        target = self.dir / "flat.pdf"

        code, result = tool("fill", str(self.form()), str(self.values({"full_name": "Erika Mustermann"})), str(target), "--flatten")

        self.assertEqual(code, 0, result)
        self.assertFalse(PdfReader(target).get_fields())
        with pdfplumber.open(target) as document:
            self.assertIn("Erika Mustermann", document.pages[0].extract_text())

    def test_fill_refuses_unknown_fields_and_wrong_options(self) -> None:
        code, result = tool("fill", str(self.form()), str(self.values({"nmae": "x"})), str(self.dir / "x.pdf"))
        self.assertEqual(code, 1)
        self.assertIn("nmae", result["error"])
        self.assertIn("full_name", result["error"])

        code, result = tool("fill", str(self.form()), str(self.values({"country": "FR", "size": "huge"})), str(self.dir / "x.pdf"))
        self.assertEqual(code, 1)
        self.assertIn("country", result["error"])
        self.assertIn("size", result["error"])
        self.assertFalse((self.dir / "x.pdf").exists())

        code, result = tool("fill", str(self.report()), str(self.values({"a": "b"})), str(self.dir / "x.pdf"))
        self.assertEqual(code, 1)
        self.assertIn("no fillable form fields", result["error"])

    # render

    def test_render_writes_one_png_per_page(self) -> None:
        out = self.dir / "png"

        code, result = tool("render", str(self.report(pages=3)), str(out), "--pages", "1,3", "--dpi", "50")

        self.assertEqual(code, 0, result)
        self.assertEqual([image["page"] for image in result["images"]], [1, 3])
        for image in result["images"]:
            with Image.open(image["png"]) as png:
                self.assertEqual(png.format, "PNG")
                self.assertAlmostEqual(png.width, round(595 * 50 / 72), delta=2)

    def test_render_ends_with_a_json_error_when_the_timeout_hits(self) -> None:
        big = self.report("big.pdf", pages=40)

        code, result = tool("render", str(big), str(self.dir / "big"), "--dpi", "600", "--timeout", "1")

        self.assertEqual(code, 1)
        self.assertIn("did not finish", result["error"])

    # failures and the round trip

    def test_a_missing_or_wrong_file_is_a_json_error(self) -> None:
        for command in ("info", "extract", "fields"):
            code, result = tool(command, str(self.dir / "nope.pdf"))
            self.assertEqual((code, "does not exist" in result["error"]), (1, True), command)
        not_a_pdf = self.dir / "plain.pdf"
        not_a_pdf.write_text("hello", encoding="utf-8")
        code, result = tool("render", str(not_a_pdf), str(self.dir / "png"))
        self.assertEqual(code, 1)
        self.assertIn("not a PDF", result["error"])
        code, result = tool("fill", str(self.form()), str(self.dir / "missing.json"), str(self.dir / "x.pdf"))
        self.assertEqual(code, 1)
        self.assertIn("does not exist", result["error"])

    def test_filled_output_opens_in_libreoffice(self) -> None:
        target = self.dir / "filled.pdf"
        code, result = tool("fill", str(self.form()), str(self.values({"full_name": "Max"})), str(target))
        self.assertEqual(code, 0, result)
        profile = self.dir / "profile"

        command = ["soffice", f"-env:UserInstallation={profile.as_uri()}", "--headless", "--convert-to", "png"]
        command += ["--outdir", str(self.dir / "lo"), str(target)]
        for _round in range(2):  # a first start on a fresh profile may only create the profile
            done = subprocess.run(
                command, check=False, capture_output=True, text=True, timeout=170, stdin=subprocess.DEVNULL
            )
            if (self.dir / "lo" / "filled.png").is_file():
                break

        self.assertTrue((self.dir / "lo" / "filled.png").is_file(), done.stdout + done.stderr)


if __name__ == "__main__":
    unittest.main()
