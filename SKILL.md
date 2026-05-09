---
name: docx2typ
description: This skill should be used when the user asks to convert a .docx file to Typst, transform a Word document into a .typ project, or extract content from a DOCX file for use in Typst writing. Handles text content, tables, images, and section structure from Chinese government reports and academic documents created with WPS Office or Microsoft Word. NOT for converting .doc (old binary) or .pdf files — those require separate preprocessing steps.
---

# docx2typ

Convert `.docx` files into a compilable Typst `.typ` file, including text content, tables, embedded images, and section hierarchy detection.

## Setup

```bash
cd ~/.claude/skills/docx2typ/scripts/docx2typ
dotnet restore
```

## Usage

```bash
cd ~/.claude/skills/docx2typ/scripts/docx2typ
dotnet run -- <input.docx>
```

Creates `{filename}-typ/` next to the input file with:

```
{filename}-typ/
├── main.typ            # Format rules + content (一体化，无需额外 import)
├── .vscode/settings.json  # Tinymist config
├── figures/            # Extracted images
└── ref.bib             # Bibliography stub
```

### Build PDF

```bash
cd {filename}-typ/
typst compile main.typ    # or open main.typ in VSCode with Tinymist
```

## Design Rationale

### Why single-file output?

Each `.docx` becomes **one** `main.typ` with formatting rules + content together. Typst's `#import` does NOT propagate `#set`/`#show` rules, so a separate `format.typ` approach doesn't work.

### Why per-level heading numbering?

The converter outputs `=` / `==` / `===` heading levels (styleId `1`→`=`, `2`→`==`, `4`→`===`). Numbering is set via a custom function:

| Level | Format | Example |
|-------|--------|---------|
| 1 | 一、 | 一、土壤类型 |
| 2 | （一） | （一）分类体系 |
| 3 | 1. | 1. 土纲 |
| 4 | (1) | (1) 亚纲 |
| 5 | a. | a. 土类 |

### Why `#table()` with flat cells?

Typst's `#table()` uses a flat list of content blocks (one per cell), unlike LaTeX's row-based `&` / `\\` syntax. This maps naturally to the DOCX traversal order. Merged cells use `#table.cell(colspan: N, rowspan: M)[...]`. The `stroke: 0.5pt` option provides Word-style grid borders.

### Why `fr` column widths?

Column weight ratios are derived from the DOCX `tblGrid` widths and expressed as `(3fr, 5fr, 2fr, ...)`. Typst's `fr` unit divides the available space proportionally — simpler and more robust than LaTeX's `X` column weight system.

### Why no `.latexmkrc` or format file?

Typst compiles with a single `typst compile` command. The Tinymist VSCode extension handles preview and auto-compilation. No build automation scripts needed.

## Section Structure Detection

Two detection paths run in parallel (identical to docx2tex):

### Path A: style-priority (styleId)

| Style ID | Mapping | Example |
|----------|---------|---------|
| `1` | `= ` | 一、工作背景 |
| `2` | `== ` | (一) 工作组织 |
| `4` | `=== ` | 1. 培肥改良土壤 |

### Path B: text fallback (no style)

- `一、` → `= `
- `（一）` → `== `
- `1.` / `1．` → `=== `

## Table Handling

All tables use Typst `#table()` with:

### Structural fidelity
- **Merged cells**: `GridSpan` → `colspan: N`, `VerticalMerge` → `rowspan: N`
- **Grid lines**: `stroke: 0.5pt` (uniform grid)
- **Header rows**: `table.header(...)` for repeat-on-page-break headers
- **Auto page break**: Typst handles table跨页 natively — no `longtblr` complexity

### Visual fidelity
- **Column widths**: `fr` units proportional to DOCX `tblGrid` widths
- **Font size**: Detected from DOCX cell runs, applied via `text(size: Npt, table(...))` wrapper — preserves both font size and page-breaking behavior
- **Inline formatting**: `#text(weight: "bold")[]`, `#text(style: "italic")[]`, `#super[]`, `#sub[]`

### Text escaping

Typst content blocks `[...]` require escaping of `\`, `#`, `[`, `]`, `<`, `>` — `<>` is interpreted as a label delimiter. Chinese characters need no escaping at all.

### Critical Typst Gotchas (from real-world testing)

1. **`table.cell()` has NO `#` prefix** — inside `table()` arguments, code mode is active; `table.cell(colspan: 2)[...]` is correct, `#table.cell(...)` fails
2. **`#figure()` blocks page breaking by default** — must add `#show figure.where(kind: "table"): set block(breakable: true)` for tables
3. **`#figure()` requires `supplement`** in Typst ≥0.14 — set `supplement: [表]` / `supplement: [图]` per call
4. **`{ set text(); table() }` code blocks prevent page breaking** — use `text(size: Npt, table(...))` wrapper instead
5. **Trailing commas inside code blocks are illegal** — `{ set text(); table(...), }` fails; comma belongs after block close `},` in the parent `#figure()` call
6. **`#sym.zws` not needed** — Typst handles digit string wrapping natively in table cells, no break-hint insertion required

## Caption Handling

### Table captions
Strips `表N.N` or `附表N` prefixes. Typst auto-numbering handles the rest. Wrapped in `#figure(table(...), caption: [...], kind: "table")`.

### Figure captions
Strips `图N(.N)(a-z)?` prefixes. Wrapped in `#figure(image(...), caption: [...], kind: "image")`.

## Image Handling

- **XML direct scan**: extracts images by scanning paragraph XML for `a:blip` references
- **TIFF auto-conversion**: TIFF images are automatically converted to PNG using `System.Drawing.Common` during extraction (Typst does not support TIFF natively)
- Images extracted to `figures/` as `figure1.jpg`, etc.
- Typst natively supports SVG, PNG, JPG

## Table Page Breaking

Typst `table()` breaks across pages automatically — no special configuration needed. Long tables with hundreds of rows will page-break naturally.

However, if a table is wrapped inside a `{ }` code block (e.g., for font-size scoping), the block prevents page breaking. The converter uses `text(size: Npt, table(...))` instead, which preserves both the font size and page-breaking behavior.

## Known Limitations

- **Numbering style**: hardcoded as `#set heading(numbering: "一、")`. Adjust manually for other styles.
- **Cell background colors**: not yet mapped from `tcPr/shd` to Typst's `fill:` option.
- **SVG images**: extracted as `.bin` — Typst supports SVG, but the DOCX's SVG content type is not yet detected. Manually rename to `.svg` if needed.
- **`.doc` (old binary)**: not supported. Convert to `.docx` first.

## Requirements

- .NET SDK 8.0+
- NuGet: `DocumentFormat.OpenXml` 3.2.0 (auto-restored)
- Typst CLI (`typst compile`) or Tinymist VSCode extension
