"""The docx skill's helper, end to end.

Needs Python 3 with python-docx, pandoc and LibreOffice (`soffice`) on PATH. SKILL_DIR names the bundle (default
skills/docx next to this repository's tests). The fixtures are built at run time, so no binary file is committed.
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

TOOL = Path(os.environ.get("SKILL_DIR", Path(__file__).resolve().parents[2] / "skills" / "docx")) / "scripts" / "docx_tool.py"


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


def tracked_run(paragraph, tag: str, text: str) -> None:
    """Append a tracked insertion (w:ins) or deletion (w:del) by "Reviewer" to the paragraph."""
    change = OxmlElement(tag)
    change.set(qn("w:id"), "90" if tag == "w:ins" else "91")
    change.set(qn("w:author"), "Reviewer")
    change.set(qn("w:date"), "2026-01-01T00:00:00Z")
    run = OxmlElement("w:r")
    content = OxmlElement("w:t" if tag == "w:ins" else "w:delText")
    content.text = text
    content.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    run.append(content)
    change.append(run)
    paragraph._p.append(change)


class DocxToolTest(unittest.TestCase):
    def setUp(self) -> None:
        self.dir = Path(tempfile.mkdtemp(prefix="docx-skill-"))

    def document(self, name: str = "in.docx", review: bool = True) -> Path:
        """A title, two headings, a styled run, a 3x2 table; with review, one comment and two tracked changes."""
        book = docx.Document()
        book.add_paragraph("Quarterly report", style="Title")
        book.add_heading("Summary", level=1)
        body = book.add_paragraph("Revenue grew ")
        body.add_run("strongly", style="Strong")
        body.add_run(" this quarter.")
        book.add_heading("Figures", level=2)
        table = book.add_table(rows=3, cols=2)
        table.style = "Table Grid"
        for row, (label, value) in enumerate((("Region", "Amount"), ("North", "5"), ("South", "7"))):
            table.cell(row, 0).text = label
            table.cell(row, 1).text = value
        book.add_paragraph("Figures are in thousands.")  # Word wants a paragraph after a table
        if review:
            note = book.add_paragraph("Margins stayed flat")
            tracked_run(note, "w:ins", " and costs fell")
            tracked_run(note, "w:del", " by luck")
            book.add_comment(body.runs[1], text="Source for this?", author="Editor")
        path = self.dir / name
        book.save(path)
        return path

    def soffice(self, source: Path, fmt: str) -> Path:
        """Convert with LibreOffice directly, not through the helper, into self.dir/lo."""
        out = self.dir / "lo"
        command = ["soffice", f"-env:UserInstallation={(self.dir / 'profile').as_uri()}", "--headless"]
        command += ["--convert-to", fmt, "--outdir", str(out), str(source)]
        produced = out / f"{source.stem}.{fmt.split(':')[0]}"
        for _round in range(2):  # a first start on a fresh profile may only create the profile
            done = subprocess.run(
                command, check=False, capture_output=True, text=True, timeout=170, stdin=subprocess.DEVNULL
            )
            if produced.is_file():
                break
        self.assertTrue(produced.is_file(), done.stderr)
        return produced

    def test_inspect_lists_styles_headings_tables_sections_changes_and_comments(self) -> None:
        code, result = tool("inspect", str(self.document()))

        self.assertEqual(code, 0, result)
        self.assertEqual(
            result["headings"],
            [
                {"level": 0, "style": "Title", "text": "Quarterly report"},
                {"level": 1, "style": "Heading 1", "text": "Summary"},
                {"level": 2, "style": "Heading 2", "text": "Figures"},
            ],
        )
        self.assertEqual(result["paragraph_styles"]["Normal"], 3 + 6, "body paragraphs and table cell paragraphs")
        self.assertEqual(result["character_styles"], {"Strong": 1})
        self.assertEqual(
            result["tables"],
            [{"index": 0, "rows": 3, "columns": 2, "style": "Table Grid", "first_row": ["Region", "Amount"]}],
        )
        section = result["sections"][0]
        self.assertEqual(section["orientation"], "portrait")
        self.assertGreater(section["page_height_mm"], section["page_width_mm"])
        changes = result["tracked_changes"]
        self.assertEqual((changes["insertions"], changes["deletions"], changes["total"]), (1, 1, 2), changes)
        self.assertEqual(changes["authors"], ["Reviewer"])
        self.assertEqual(result["comments"], [{"author": "Editor", "text": "Source for this?"}])

    def test_inspect_reports_no_review_marks_on_a_clean_document(self) -> None:
        code, result = tool("inspect", str(self.document(review=False)))

        self.assertEqual(code, 0, result)
        self.assertEqual((result["tracked_changes"]["total"], result["comment_count"]), (0, 0))

    def test_inspect_rejects_a_missing_or_wrong_file(self) -> None:
        code, result = tool("inspect", str(self.dir / "nope.docx"))
        self.assertEqual((code, "does not exist" in result["error"]), (1, True))
        plain = self.dir / "plain.docx"
        plain.write_text("hello", encoding="utf-8")
        code, result = tool("inspect", str(plain))
        self.assertEqual(code, 1)
        self.assertIn("convert", result["error"])

    def test_text_gives_markdown_with_changes_accepted_by_default(self) -> None:
        code, result = tool("text", str(self.document()))

        self.assertEqual(code, 0, result)
        markdown = result["markdown"]
        self.assertIn("title: Quarterly report", markdown, "a Title paragraph comes back as front matter")
        self.assertIn("# Summary", markdown)
        self.assertIn("## Figures", markdown)
        self.assertIn("South", markdown)
        self.assertIn("and costs fell", markdown)
        self.assertNotIn("by luck", markdown)

    def test_text_can_reject_or_show_all_changes(self) -> None:
        source = self.document()
        code, rejected = tool("text", str(source), "--changes", "reject")
        self.assertEqual(code, 0, rejected)
        self.assertIn("by luck", rejected["markdown"])
        self.assertNotIn("and costs fell", rejected["markdown"])

        code, shown = tool("text", str(source), "--changes", "all")

        self.assertEqual(code, 0, shown)
        self.assertIn(".insertion", shown["markdown"])
        self.assertIn(".deletion", shown["markdown"])
        self.assertIn("Source for this?", shown["markdown"])

    def test_text_writes_to_a_file_on_request(self) -> None:
        target = self.dir / "out" / "report.md"

        code, result = tool("text", str(self.document()), "--output", str(target))

        self.assertEqual(code, 0, result)
        self.assertNotIn("markdown", result)
        self.assertIn("# Summary", target.read_text(encoding="utf-8"))

    def test_text_rejects_a_doc_and_a_missing_file(self) -> None:
        legacy = self.dir / "old.doc"
        legacy.write_bytes(b"\xd0\xcf\x11\xe0")
        code, result = tool("text", str(legacy))
        self.assertEqual(code, 1)
        self.assertIn("convert it to .docx first", result["error"])
        code, result = tool("text", str(self.dir / "nope.docx"))
        self.assertEqual((code, "does not exist" in result["error"]), (1, True))

    def test_convert_docx_to_pdf(self) -> None:
        target = self.dir / "out" / "report.pdf"

        code, result = tool("convert", str(self.document()), str(target))

        self.assertEqual(code, 0, result)
        self.assertEqual(target.read_bytes()[:5], b"%PDF-")
        words = subprocess.run(["pdftotext", str(target), "-"], check=True, capture_output=True, text=True).stdout
        self.assertIn("Quarterly report", words)

    def test_convert_doc_and_odt_to_docx_keeps_text_and_headings(self) -> None:
        source = self.document(review=False)
        legacy = self.soffice(source, 'doc:MS Word 97')
        code, result = tool("convert", str(source), str(self.dir / "open.odt"))
        self.assertEqual(code, 0, result)

        for given in (legacy, self.dir / "open.odt"):
            target = self.dir / f"from-{given.suffix[1:]}.docx"
            code, result = tool("convert", str(given), str(target))
            self.assertEqual(code, 0, result)
            code, result = tool("inspect", str(target))
            self.assertEqual(code, 0, result)
            self.assertIn({"level": 1, "style": "Heading 1", "text": "Summary"}, result["headings"], given)
            self.assertEqual(result["tables"][0]["first_row"], ["Region", "Amount"], given)

    def test_convert_rejects_an_unknown_target_and_a_missing_input(self) -> None:
        code, result = tool("convert", str(self.document()), str(self.dir / "out.png"))
        self.assertEqual(code, 1)
        self.assertIn(".docx, .pdf, .odt", result["error"])
        code, result = tool("convert", str(self.dir / "nope.doc"), str(self.dir / "out.docx"))
        self.assertEqual((code, "does not exist" in result["error"]), (1, True))

    def test_convert_uses_a_private_profile_so_parallel_calls_succeed(self) -> None:
        def one(index: int) -> tuple[int, dict]:
            return tool("convert", str(self.document(f"p{index}.docx")), str(self.dir / f"p{index}.pdf"))

        with ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(one, (1, 2)))

        for code, result in results:
            self.assertEqual(code, 0, result)

    def test_convert_ends_with_a_json_error_when_the_timeout_hits(self) -> None:
        book = docx.Document()
        for index in range(20000):
            book.add_paragraph(f"Paragraph {index} " + "filler text " * 20)
        big = self.dir / "big.docx"
        book.save(big)

        code, result = tool("convert", str(big), str(self.dir / "big.pdf"), "--timeout", "1")

        self.assertEqual(code, 1)
        self.assertIn("did not finish", result["error"])

    def test_output_opens_in_libreoffice(self) -> None:
        target = self.dir / "round.docx"
        code, result = tool("convert", str(self.document()), str(target))
        self.assertEqual(code, 0, result)

        plain = self.soffice(target, "txt:Text")

        self.assertIn("Quarterly report", plain.read_text(encoding="utf-8-sig"))


if __name__ == "__main__":
    unittest.main()
