---
name: pptx
description: Use when the task reads, builds, edits or converts a presentation — a PowerPoint deck (.pptx), a .ppt or .odp file — or turns content into slides. Covers reading a deck's text, building slides on the deck's own layouts, rendering every slide to an image to look at it, and checking for empty placeholders and overflowing text before the file is handed over.
---

# PowerPoint decks

Files you were given are under `inputs/<id>/`. Write every file you deliver to `outputs/`.
Python with `python-pptx` and the `markitdown` command are installed; nothing can be installed.

The helper `/skills/pptx/scripts/pptx_tool.py` has two commands and prints JSON:

```
python3 /skills/pptx/scripts/pptx_tool.py inspect DECK.pptx          # size, layouts, placeholders, text per slide
python3 /skills/pptx/scripts/pptx_tool.py render DECK OUTDIR         # OUTDIR/<name>.pdf + OUTDIR/slide-<n>.png
```

Run it with `--help` for the options (`render --dpi`, `render --timeout`).

## Look before you change

1. For the words, `markitdown DECK.pptx` prints every slide's text (and tables) as Markdown, slide by slide.
2. For the structure, `inspect` it: slide size, every layout the deck offers with its placeholders, and per
   slide the layout used, the title, which placeholders are filled or empty, all text, the number of
   pictures, tables and charts, and the speaker notes.
3. For the look, `render` it and look at the PNGs. A `.ppt` or `.odp` renders too, but `inspect` and
   python-pptx read only `.pptx`.

## Build on the deck's own layouts

- **Start from the user's template or deck** when there is one: `deck = Presentation("inputs/<id>/template.pptx")`.
  Its masters carry the fonts, colours and logo; slides you add inherit them. Without one,
  `Presentation()` gives a plain 4:3 (10 x 7.5 in) deck whose layouts are laid out for that size.
- **Pick a layout by name**, as `inspect` lists them:
  `slide = deck.slides.add_slide(next(l for l in deck.slide_layouts if l.name == "Title and Content"))`.
- **Fill placeholders instead of drawing text boxes**: `slide.shapes.title.text = "..."`,
  `slide.placeholders[1].text_frame` for the body (the idx comes from `inspect`). Text in placeholders
  follows the template, shows in the outline view and moves with a layout change. A placeholder you leave
  empty shows "Click to add text" in PowerPoint: fill it or remove it
  (`ph.element.getparent().remove(ph.element)`).
- **Remove a template's sample slides** before adding your own:

  ```python
  ids = deck.slides._sldIdLst
  for sld in list(ids):          # or only the ones you do not keep
      deck.part.drop_rel(sld.rId)
      ids.remove(sld)
  ```

- **One idea per slide.** The title states the point ("Costs fell 12 % in Q3", not "Costs"); the body
  backs it with at most about six short bullets. More than that is two slides. Detail and sources go into
  the speaker notes: `slide.notes_slide.notes_text_frame.text = "..."`.
- **Bullets are paragraphs**: `tf.text = "first"`, then `p = tf.add_paragraph(); p.text = "second"`;
  `p.level = 1` indents one step.
- **Pictures**: into a picture placeholder with `ph.insert_picture(path)` (it crops to fill), or
  `slide.shapes.add_picture(path, left, top, width=...)` with only the width given, so the aspect ratio stays.
- **Tables and charts**: `slide.shapes.add_table(rows, cols, left, top, width, height)`; a chart from
  `pptx.chart.data.CategoryChartData` with `slide.shapes.add_chart(...)` stays editable in PowerPoint, an
  image of a chart does not.
- Sizes and positions take `pptx.util.Inches`, `Cm` or `Pt`. Keep a margin of about 0.5 in to every edge.
- python-pptx does not shrink text to fit: too much text runs out of its box. Shorten it or split the slide;
  set a smaller `run.font.size` only as a last step, and not below about 14 pt for body text.

## Edit an existing deck in place

Open it, change what was asked, save under a new name in `outputs/`. Slides, layouts and formatting you did
not touch stay as they were. Assigning `text_frame.text` or `paragraph.text` replaces the runs and with them
their bold, colour and size; to keep the formatting, change the text of the existing run:
`paragraph.runs[0].text = "new"` (and clear the other runs of that paragraph if the old text was split).

## Render and look before you hand over

1. `inspect` the result: the slide count is right, every slide has its title, no placeholder is `"empty": true`.
2. `render` it: `python3 /skills/pptx/scripts/pptx_tool.py render outputs/deck.pptx work/render`. Then look at
   **every** PNG, not a sample. Look for text that runs out of its box or off the slide, shapes that overlap,
   text that is too small to read, poor contrast, stretched pictures, empty areas left from a removed shape.
3. Fix the cause, render again, look again. Do not deliver a deck you have not looked at after the last change.
4. The render uses the fonts installed here (Liberation, DejaVu); a font that is missing is replaced by one of
   these, so line breaks can differ slightly from PowerPoint. Leave some room in every text box for that.

`render` also leaves `<name>.pdf` in OUTDIR: when the user wants a PDF of the deck, that is it (copy it to
`outputs/`). A render that takes too long ends with an error; `--timeout 600` allows longer for a big deck.

## Hand-over

Tell the user which file is the result, how many slides it holds, which template or layouts it uses, and that
you rendered and looked at every slide. Say what you left in the speaker notes. If something could not be
done (a `.ppt` that has to be edited, a font that is not here, a chart type python-pptx cannot draw), say what
and why instead of delivering it silently.
