"""The pptx skill's helper, end to end.

Needs Python 3 with python-pptx, pypdf and Pillow, LibreOffice (`soffice`) and pdftoppm on PATH. SKILL_DIR names
the bundle (default skills/pptx next to this repository's tests). The fixtures are built at run time, so no binary
file is committed.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Inches
from pypdf import PdfReader

TOOL = Path(os.environ.get("SKILL_DIR", Path(__file__).resolve().parents[2] / "skills" / "pptx")) / "scripts" / "pptx_tool.py"


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


def layout(deck, name: str):
    return next(item for item in deck.slide_layouts if item.name == name)


class PptxToolTest(unittest.TestCase):
    def setUp(self) -> None:
        self.dir = Path(tempfile.mkdtemp(prefix="pptx-skill-"))

    def tearDown(self) -> None:
        shutil.rmtree(self.dir, ignore_errors=True)

    def deck(self, slides: int = 3, name: str = "deck.pptx") -> Path:
        """`slides` "Title and Content" slides, each with a title and two bullets."""
        deck = Presentation()
        for number in range(1, slides + 1):
            slide = deck.slides.add_slide(layout(deck, "Title and Content"))
            slide.shapes.title.text = f"Slide {number}"
            body = slide.placeholders[1].text_frame
            body.text = "First point"
            body.add_paragraph().text = "Second point"
        path = self.dir / name
        deck.save(path)
        return path

    def test_inspect_lists_layouts_placeholders_text_and_objects(self) -> None:
        picture = self.dir / "dot.png"
        Image.new("RGB", (40, 40), "red").save(picture)
        deck = Presentation()
        cover = deck.slides.add_slide(layout(deck, "Title Slide"))
        cover.shapes.title.text = "Quarterly review"  # subtitle left empty on purpose
        body = deck.slides.add_slide(layout(deck, "Title Only"))
        body.shapes.title.text = "Numbers"
        table = body.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(4), Inches(1)).table
        table.cell(0, 0).text, table.cell(0, 1).text = "Region", "Amount"
        table.cell(1, 0).text, table.cell(1, 1).text = "North", "5"
        body.shapes.add_picture(str(picture), Inches(6), Inches(2))
        group = body.shapes.add_group_shape()
        group.shapes.add_textbox(Inches(1), Inches(5), Inches(3), Inches(1)).text_frame.text = "Inside a group"
        body.notes_slide.notes_text_frame.text = "Say the total out loud"
        path = self.dir / "inspect.pptx"
        deck.save(path)

        code, result = tool("inspect", str(path))

        self.assertEqual(code, 0, result)
        self.assertEqual((result["slide_width_in"], result["slide_height_in"], result["slide_count"]), (10.0, 7.5, 2))
        self.assertIn("Title and Content", [item["name"] for item in result["layouts"]])
        first, second = result["slides"]
        self.assertEqual((first["layout"], first["title"]), ("Title Slide", "Quarterly review"))
        self.assertEqual(
            [(ph["type"], ph["empty"]) for ph in first["placeholders"]],
            [("CENTER_TITLE", False), ("SUBTITLE", True)],
        )
        self.assertEqual(second["text"], ["Numbers", "Region | Amount", "North | 5", "Inside a group"])
        self.assertEqual((second["tables"], second["pictures"], second["charts"]), (1, 1, 0))
        self.assertEqual(second["notes"], "Say the total out loud")
        self.assertEqual(first["notes"], "")

    def test_inspect_rejects_a_missing_or_wrong_file(self) -> None:
        code, result = tool("inspect", str(self.dir / "nope.pptx"))
        self.assertEqual((code, "does not exist" in result["error"]), (1, True))
        plain = self.dir / "plain.pptx"
        plain.write_text("hello", encoding="utf-8")
        code, result = tool("inspect", str(plain))
        self.assertEqual(code, 1)
        self.assertIn("not a .pptx", result["error"])

    def test_render_writes_a_pdf_and_one_png_per_slide_in_order(self) -> None:
        out = self.dir / "render"

        code, result = tool("render", str(self.deck(slides=11)), str(out))

        self.assertEqual(code, 0, result)
        self.assertEqual(result["slides"], 11)
        self.assertEqual([Path(p).name for p in result["images"]][:3], ["slide-01.png", "slide-02.png", "slide-03.png"])
        self.assertEqual(Path(result["images"][-1]).name, "slide-11.png")
        self.assertEqual(len(PdfReader(result["pdf"]).pages), 11)
        with Image.open(result["images"][0]) as image:
            self.assertEqual(image.size, (1000, 750))  # 10 x 7.5 in at the default 100 dpi

    def test_render_again_drops_pngs_of_slides_that_are_gone(self) -> None:
        out = self.dir / "render"
        code, result = tool("render", str(self.deck(slides=3)), str(out))
        self.assertEqual(code, 0, result)
        (out / "slide-notes.png").write_bytes(b"not ours")

        code, result = tool("render", str(self.deck(slides=1, name="short.pptx")), str(out), "--dpi", "50")

        self.assertEqual(code, 0, result)
        self.assertEqual(sorted(p.name for p in out.glob("slide-*.png")), ["slide-1.png", "slide-notes.png"])

    def test_render_reports_a_missing_file(self) -> None:
        code, result = tool("render", str(self.dir / "nope.pptx"), str(self.dir / "out"))

        self.assertEqual(code, 1)
        self.assertIn("does not exist", result["error"])

    def test_render_uses_a_private_profile_so_parallel_calls_succeed(self) -> None:
        def one(index: int) -> tuple[int, dict]:
            return tool("render", str(self.deck(name=f"p{index}.pptx")), str(self.dir / f"out{index}"))

        with ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(one, (1, 2)))

        for code, result in results:
            self.assertEqual(code, 0, result)
            self.assertEqual(result["slides"], 3)

    def test_render_ends_with_a_json_error_when_the_timeout_hits(self) -> None:
        code, result = tool("render", str(self.deck(slides=300)), str(self.dir / "big"), "--timeout", "1")

        self.assertEqual(code, 1)
        self.assertIn("did not finish", result["error"])

    def test_every_subcommand_has_help(self) -> None:
        for command in ("inspect", "render"):
            done = subprocess.run(
                [sys.executable, "-I", str(TOOL), command, "--help"], check=False, capture_output=True, text=True
            )
            self.assertEqual(done.returncode, 0, done.stderr)
            self.assertIn("usage:", done.stdout)

    def test_markitdown_reads_slide_text(self) -> None:
        done = subprocess.run(
            ["markitdown", str(self.deck(slides=2))], check=False, capture_output=True, text=True, timeout=120
        )

        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertIn("Slide 2", done.stdout)
        self.assertIn("Second point", done.stdout)

    def test_a_deck_built_from_a_template_opens_in_libreoffice(self) -> None:
        """The SKILL.md recipe: open the template, drop its sample slides, add slides on its own layouts."""
        template = self.deck(slides=2, name="template.pptx")
        deck = Presentation(str(template))
        slide_ids = deck.slides._sldIdLst
        for slide_id in list(slide_ids):
            deck.part.drop_rel(slide_id.rId)
            slide_ids.remove(slide_id)
        cover = deck.slides.add_slide(layout(deck, "Title Slide"))
        cover.shapes.title.text = "Built from the template"
        cover.placeholders[1].text = "Subtitle"
        target = self.dir / "built.pptx"
        deck.save(target)
        code, result = tool("inspect", str(target))
        self.assertEqual((code, result["slide_count"]), (0, 1), result)
        self.assertFalse(any(ph["empty"] for ph in result["slides"][0]["placeholders"]))

        profile, out = self.dir / "profile", self.dir / "lo"
        command = ["soffice", f"-env:UserInstallation={profile.as_uri()}", "--headless", "--convert-to", "pdf"]
        command += ["--outdir", str(out), str(target)]
        for _round in range(2):  # a first start on a fresh profile may only create the profile
            done = subprocess.run(
                command, check=False, capture_output=True, text=True, timeout=170, stdin=subprocess.DEVNULL
            )
            if (out / "built.pdf").is_file():
                break

        self.assertTrue((out / "built.pdf").is_file(), done.stderr)
        reader = PdfReader(out / "built.pdf")
        self.assertEqual(len(reader.pages), 1)
        self.assertIn("Built from the template", reader.pages[0].extract_text())


if __name__ == "__main__":
    unittest.main()
