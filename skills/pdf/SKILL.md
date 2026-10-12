---
name: pdf
description: Use when the task reads, builds, edits or converts a PDF — pulling text or tables out of it, reading a scan, filling a form, merging, splitting, rotating or unlocking pages, or producing a new PDF. Covers telling a scan from a text PDF and OCRing it, filling AcroForm fields by name, and rendering pages to look at them before the file is handed over.
---

# PDF files

Files you were given are under `inputs/<id>/`. Write every file you deliver to `outputs/`.
Python with `pypdf`, `pdfplumber` and `reportlab` is installed, and the programs `qpdf`, `pdftoppm` and
`tesseract` (German and English); nothing can be installed.

The helper `/skills/pdf/scripts/pdf_tool.py` has five commands and prints JSON:

```
python3 /skills/pdf/scripts/pdf_tool.py info FILE.pdf                    # pages, text layer per page, fields, encryption
python3 /skills/pdf/scripts/pdf_tool.py extract FILE.pdf [--ocr]         # text and tables per page
python3 /skills/pdf/scripts/pdf_tool.py fields FILE.pdf                  # fillable form fields
python3 /skills/pdf/scripts/pdf_tool.py fill FILE.pdf VALUES.json OUT.pdf [--flatten]
python3 /skills/pdf/scripts/pdf_tool.py render FILE.pdf OUTDIR [--pages 1-3]   # one PNG per page
```

Every command takes `--password` for an encrypted file; `extract` and `render` take `--pages 1-3,5`.
Run it with `--help` for the other options.

## Look before you read

1. `info` the file first. `pages_without_text` lists the pages that have no text layer: those are scans
   (or pure pictures), and plain text extraction returns nothing for them. `locked: true` means it needs
   a password; ask the user for it rather than guessing.
2. `extract` gives each page's text and the tables pdfplumber finds (a table is a list of rows, each a list
   of cells). Add `--ocr` when `info` showed pages without text: those pages are read by OCR in German and
   English (`--lang eng` for English only), and `"source": "ocr"` marks them. A scan can carry a little text
   (a stamped page number, a scanner's header line) and then does not count as without text: when a page's
   `text` is far shorter than the page looks, render it, look at the PNG, and if it is a scan, read it with
   `--ocr-pages 3,5-6`, which OCRs those pages whatever text they have.
3. OCR is a reading, not the document. Numbers, names and amounts from an OCR page can be wrong (a `1` read
   as `/`, `0` as `O`): render that page and look at the PNG before you rely on a figure, and tell the user
   which values came from OCR.
4. For finer work than `extract`, open the file with `pdfplumber` yourself: `page.extract_words()` for
   positions, `page.crop(bbox)` for one region, `page.extract_table(table_settings=...)` for a table
   without ruled lines.

## Change pages with qpdf

`qpdf` keeps everything it does not touch (fonts, links, form fields) and never re-renders a page:

```
qpdf --empty --pages a.pdf b.pdf -- outputs/merged.pdf          # merge in this order
qpdf in.pdf --pages . 2-4 -- outputs/part.pdf                   # keep pages 2 to 4
qpdf in.pdf --split-pages outputs/page.pdf                      # one file per page: page-01.pdf, page-02.pdf, ...
qpdf in.pdf outputs/rotated.pdf --rotate=+90:1,3                # turn pages 1 and 3 clockwise
qpdf --password=PW --decrypt in.pdf outputs/open.pdf            # remove the password (only with the owner's say)
```

`pypdf` does the same from Python (`PdfWriter`, `writer.append(...)`, `page.rotate(90)`) when you need
logic in between.

## Fill a form

1. `fields` lists every field with its exact `name`, `type` (text, checkbox, radio, choice), current value,
   allowed `options` and page. Use those names as they are; a label printed on the page is not the name.
2. Write the values to a JSON file, `{"full_name": "Erika Mustermann", "agree": true, "size": "large"}`:
   text as a string, a checkbox as `true`/`false`, a radio or choice as one of its options.
3. `fill FILE.pdf values.json outputs/filled.pdf`. An unknown name or a value that is not an option stops
   the command with a list of what the file has; nothing is written then.
4. `--flatten` draws the values into the page and removes the form, so nobody can change them later. Only
   do it when the user asks for a final, non-editable copy.

A PDF without fields (`fields` is empty) is not a form: build an overlay with reportlab, put the text at the
right coordinates, and merge it onto the page with `page.merge_page(overlay_page)` in pypdf.

## Build a new PDF

Use reportlab: `SimpleDocTemplate` with `Paragraph`, `Table` and `Spacer` for a document that flows over
pages, `canvas.Canvas` when you place things by coordinates (points, origin bottom left, A4 is 595 x 842).
The built-in fonts cover Western European text, umlauts and `€` included. For anything else, register
a TrueType font first:

```python
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont("DejaVuSans", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
```

Set a title (`SimpleDocTemplate(path, title=...)`), keep margins of at least 2 cm, and let tables repeat
their header row on each page (`Table(rows, repeatRows=1)`).

## Check before you hand over

1. `render` every page you made or changed (`--pages` for a long file) and look at the PNGs: text cut off,
   overlapping, a table running off the page, a field value in the wrong box, a page upside down.
   `--dpi 150` helps for small print.
2. `info` the result: the page count you expect, and `fields` still there unless you flattened.
3. `extract` the result and check that the text you put in reads back (a missing glyph shows up as a gap
   or a box here and on the PNG).
4. Fix the cause and check again. Do not deliver a file you have not looked at.

## Hand-over

Tell the user which file is the result, how many pages it has, what you changed or filled, and what you
checked. Name the values that came from OCR and any page you could not read. If you could not make something
work, say what and why instead of delivering it silently.
