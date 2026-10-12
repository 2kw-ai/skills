# 2kw skills

Skills that teach an agent to do a job well. The first ones cover office files: Excel, Word, PDF and
PowerPoint, each with a tested helper that checks the result before it is handed over.

| Skill | What it does |
|---|---|
| `xlsx` | Reads, builds, edits and converts Excel workbooks (`.xlsx`, `.csv`, `.xls`); recalculates and checks them before delivery |
| `docx` | Reads Word documents as Markdown, edits them in place keeping their styles, builds new ones, converts `.doc`/`.odt` and to PDF |
| `pdf` | Extracts text and tables, OCRs scans (German, English), fills forms, merges, splits, rotates and unlocks pages, creates PDFs |
| `pptx` | Reads and builds slide decks with the deck's own layouts; renders every slide to an image to check it before delivery |

## Use on 2kw.ai

1. **Skills → Install plugin**, repository `https://github.com/2kw-ai/skills`.
2. Bind the skills you want (`xlsx`, `docx`, `pdf`, `pptx`) to an agent that has the **Bash** tool on.

## Use in Claude Code

```
/plugin marketplace add 2kw-ai/skills
/plugin install skills@2kw-skills
```

## Requirements

Python 3 with `openpyxl`, `python-docx`, `python-pptx`, `pypdf`, `pdfplumber`, `reportlab` and `markitdown`;
LibreOffice (`soffice`), pandoc, qpdf, poppler (`pdftoppm`) and tesseract (`deu`, `eng`). Nothing is
installed at run time.

## Tests

```
python -m unittest discover -s tests/<skill> -v
```

Needs the requirements above on PATH. `SKILL_DIR=<path>` tests a bundle somewhere else.
