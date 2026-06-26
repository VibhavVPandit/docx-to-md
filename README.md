# docx → md

Convert Word (`.docx`) documents to Markdown (`.md`) while preserving as much
structure and inline formatting as possible.

## Features

Preserved:

- **Bold**, *italic*, ***bold italic***, and ~~strikethrough~~
- Headings H1–H6 (`Heading 1`…`Heading 6`, plus `Title`/`Subtitle`)
- Bulleted and numbered lists, including nesting (ordered vs. bulleted is
  detected from the document's numbering definitions)
- Hyperlinks → `[text](url)` (external links and internal anchors)
- Inline code / monospace text → `` `code` `` (detected by font; auto-disabled
  for documents that are *entirely* monospace — see below)
- Tables as Markdown pipe tables, with bold/italic preserved inside cells
- Tables with **merged or nested cells** → rendered as raw HTML (`colspan` /
  `rowspan`), which Markdown permits
- Images → extracted to an `images/` folder and referenced as `![alt](path)`
- Paragraph breaks, soft line breaks, and horizontal rules

Markdown special characters (`` \ ` * _ [ ] ``, and `|` inside table cells) that
appear literally in the source are escaped so they render as text.

## Install

```bash
python -m venv mdconvert
source mdconvert/Scripts/activate     # Windows (Git Bash)
# mdconvert\Scripts\activate          # Windows (PowerShell/cmd)
# source mdconvert/bin/activate       # macOS / Linux
pip install -r requirements.txt
```

## Usage

```bash
python docx_to_md.py input.docx -o output.md
```

Options:

| Flag | Description |
| --- | --- |
| `-o`, `--output` | Output path (default: input name with `.md`). |
| `--images-dir DIR` | Folder for extracted images, relative to the output file (default: `images`). |
| `--html-tables` | Render *every* table as HTML, not just merged/nested ones. |
| `--code-detection auto\|on\|off` | Whether monospace-font text becomes inline `` `code` ``. `auto` (default) keeps it for normal prose but disables it when the document is predominantly monospace; `on` forces it; `off` disables it entirely. |
| `--engine custom\|mammoth` | Conversion engine. `custom` (default) is the python-docx walker; `mammoth` is a quick baseline used for comparison. |
| `--verbose` | Print warnings and progress to stderr. |

Examples:

```bash
python docx_to_md.py report.docx                 # -> report.md (+ images/)
python docx_to_md.py report.docx -o out/r.md --images-dir assets --verbose
python docx_to_md.py report.docx --html-tables   # all tables as HTML
python docx_to_md.py report.docx --engine mammoth -o baseline.md
```

## Behaviour notes & known limitations

- **`.docx` only.** Legacy binary `.doc` files are not supported — convert them
  to `.docx` first (e.g. with LibreOffice/Word).
- **Footnotes, comments, and tracked changes are dropped.** The final accepted
  text is kept: tracked insertions are accepted and tracked deletions removed.
- **Color, font family, and font size are lost** — Markdown cannot represent
  them. Monospace fonts are the one exception (mapped to inline code) — except
  when the *whole* document is monospace (e.g. a log or terminal dump), where
  that would wrap everything in backticks and hide bold/italic. In that case
  code detection is turned off automatically; override with `--code-detection`.
- **Underline is not preserved** — Markdown has no underline; the text remains
  but the underline is dropped.
- **Merged / nested tables become HTML.** Standard Markdown pipe tables cannot
  express `rowspan`/`colspan` or nested tables, so those render as an HTML
  `<table>` (with a `<!-- note: ... -->` marker). Use `--html-tables` to force
  this for all tables.
- Anything the converter cannot represent produces a warning (see `--verbose`)
  and, where applicable, an HTML comment in the output so nothing is silently
  lost.

## Project layout

```
docx_to_md.py            # converter + CLI
requirements.txt         # pinned dependencies
README.md
```
