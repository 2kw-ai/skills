---
name: docx
description: Use when the task reads, builds, edits or converts a Word document — a .docx, a .doc or an .odt file — or turns one into a PDF. Covers reading it as Markdown, writing a new one with real styles, changing an existing one in place without losing its formatting, tracked changes and comments, and checking the result before it is handed over.
---

# Word documents

Files you were given are under `inputs/<id>/`. Write every file you deliver to `outputs/`.
Python with `python-docx` is installed, and so are pandoc and LibreOffice; nothing can be installed.

The helper `/skills/docx/scripts/docx_tool.py` has three commands and prints JSON:

```
python3 /skills/docx/scripts/docx_tool.py inspect FILE.docx      # styles in use, headings, tables, sections, changes, comments
python3 /skills/docx/scripts/docx_tool.py text FILE.docx         # the document as Markdown (pandoc)
python3 /skills/docx/scripts/docx_tool.py convert IN OUT         # LibreOffice; OUT's suffix: .docx, .pdf or .odt
```

Run it with `--help` for the options.

## Read it first

1. A `.doc` (or `.rtf`) is converted before anything else: `convert inputs/1/old.doc work/old.docx`.
   python-docx and pandoc read `.docx`; pandoc also reads `.odt`.
2. `text FILE` gives the content as Markdown: headings as `#`, tables, lists, emphasis. A `Title`
   paragraph comes back as `title:` in the front matter. Long documents: `--output work/doc.md`.
3. `inspect FILE` tells you how the document is built: which paragraph and character styles it uses and
   how often, the headings, each table's size and first row, page size and margins per section, and how
   many tracked changes and comments it carries (with authors).
4. Tracked changes: `text` accepts them by default. `--changes reject` shows the text before them,
   `--changes all` marks each insertion, deletion and comment so you can report who changed what.

## Change an existing document in place

The user's file has its own styles, numbering, headers, footers, fields and page setup. Keep them:
open the file, change only what the task asks for, save under a new name in `outputs/`.

```python
import docx
doc = docx.Document("inputs/1/contract.docx")
# ... edit ...
doc.save("outputs/contract.docx")
```

- **Never rebuild it.** Creating a new document and copying the text across loses everything `inspect`
  does not show you. If a change seems to need that, say so instead.
- **Keep the runs.** `paragraph.text = "..."` throws away every run's formatting (bold, links, fonts).
  Change text inside the runs; this replaces a phrase even when Word split it across several runs:

  ```python
  def replace(paragraph, old, new):
      """Replace old with new in one paragraph; new takes the formatting of the run where old starts."""
      runs = paragraph.runs
      start = "".join(run.text for run in runs).find(old)
      if start < 0:
          return False
      end, a = start + len(old), 0
      for run in runs:
          b = a + len(run.text)
          if a < end and b > start:
              first = new if a <= start else ""
              run.text = run.text[: max(start - a, 0)] + first + run.text[min(end, b) - a :]
          a = b
      return True
  ```

  Walk `doc.paragraphs`, the cells of `doc.tables`, and `section.header` / `section.footer` paragraphs:
  text in tables and headers is not in `doc.paragraphs`.
- **Use the document's own styles.** New paragraphs take a style name from `inspect`
  (`next_paragraph.insert_paragraph_before("text", style="Body Text")`), never direct font settings. A
  German document names its styles differently in the file; `inspect` shows what python-docx calls them.
- **Tracked changes and comments.** python-docx does not see text inside a tracked insertion, and its
  edits are not recorded as changes. If `inspect` reports open changes, tell the user before you edit.
  To mark what you changed for a reviewer, add a comment on the run:
  `doc.add_comment(run, text="Updated the amount to 150 USD", author="Assistant")`.

## Build a new document

Start from `docx.Document()` and work with styles, not with formatting on single runs:

- Structure with `doc.add_heading(text, level=1)` (levels 1 to 3), `doc.add_paragraph(text)`,
  `style="List Bullet"` / `"List Number"` for lists, `doc.add_page_break()`.
- Change the look once, on the style: `doc.styles["Normal"].font.name = "Calibri"`,
  `doc.styles["Normal"].font.size = Pt(11)` (`from docx.shared import Pt, Cm`).
- Page: the default is US Letter. For A4 set `section.page_width = Cm(21)`, `section.page_height = Cm(29.7)`
  on `doc.sections[0]`, and margins (`section.left_margin = Cm(2.5)`).
- Tables: `table = doc.add_table(rows=1, cols=3)`, `table.style = "Table Grid"`, the header in the first
  row (bold its runs), then `table.add_row()` per record. Numbers right-aligned, the same decimals throughout.
- Pictures: `doc.add_picture("chart.png", width=Cm(15))`. Header and footer text:
  `doc.sections[0].header.paragraphs[0].text = "..."`.
- Metadata: `doc.core_properties.title` and `.author`.

## Other formats

- `.odt` out: `convert outputs/report.docx outputs/report.odt`. A `.doc` cannot be written; deliver
  `.docx` and say so.
- PDF: `convert outputs/report.docx outputs/report.pdf`. LibreOffice lays it out; fonts the document asks
  for but the system lacks are substituted, so look at the result.

## Check before you hand over

1. `inspect` the result: the headings, tables and styles you expect, no tracked changes you did not mean
   to leave, page size right.
2. `text` the result and read it: no placeholder left, nothing doubled, no text lost from an edited passage.
3. Look at the layout: `convert` it to PDF in a scratch directory, render the pages with
   `pdftoppm -png -r 60 work/report.pdf work/page`, and look at the PNGs. Fix overflowing tables, empty
   pages, wrong page size. `convert` succeeding also shows that the file opens.
4. If `convert` answers with an error, read it: a missing or damaged input, or a document too large for the
   time limit (`--timeout 300` allows longer).

## Hand-over

Tell the user which file is the result, what you changed (or what the new document holds), which styles
and page setup it uses, and what you checked. Name tracked changes or comments the file still carries. If you
could not make something work, say what and why instead of delivering it silently.
