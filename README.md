# 2kw skills

Skills that teach an agent to do a job well. The first ones cover office files: formulas that stay
formulas and are calculated, checked for error values before the file is handed over.

| Skill | What it does |
|---|---|
| `xlsx` | Reads, builds, edits and converts Excel workbooks (`.xlsx`, `.csv`, `.xls`); recalculates and checks them before delivery |

## Use on 2kw.ai

1. **Skills → Install plugin**, repository `https://github.com/2kw-ai/skills`.
2. Bind `xlsx` to an agent that has the **Bash** tool on.

## Use in Claude Code

```
/plugin marketplace add 2kw-ai/skills
/plugin install skills@2kw-skills
```

## Requirements

Each skill's helper runs with Python 3 and `openpyxl`, and calls LibreOffice (`soffice`) to calculate
formulas. Nothing is installed at run time.

## Tests

```
python -m unittest discover -s tests/xlsx -v
```

Needs the requirements above on PATH. `SKILL_DIR=<path>` tests a bundle somewhere else.
