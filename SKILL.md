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

### Full Pipeline (3 steps, subagent orchestration)

To keep the main conversation context lean, each step is delegated to a **subagent** via the Agent tool. Only summaries return to the main context.

**Pre-step A**: Determine the chapter prefix and output path.

- Chapter prefix: extract from filename (`第X章` → `0X`, `ch07` → `07`). If the filename has no chapter pattern, ask the user or pass `--prefix=XX` to scripts.
- Output path: `{docx_dir}/{safe_name}-typ/main.typ` where `safe_name` replaces `[^\w\- ]` with `_`.

---

**Pre-step B**: Semantic heading extraction and mapping (**new**)

DOCX style IDs are often unreliable (custom styles, af0, etc.). Instead of relying on them, extract headings by text pattern first, then build a heading-level mapping.

**Spawn as a subagent** with `subagent_type: "general-purpose"`. Prompt:

```
Run the heading-mapping pipeline on a DOCX:

1. DOCX XML 中提取所有段落实文，识别标题模式：
   - `第X章` / `第X节` → level 0 (chapter)
   - `一、` / `二、` / `三、` ... → level 1 (section)
   - `（一）` / `（二）` / `（三）` ... → level 2 (subsection)
   - `1.` / `1、` / `2.` / `2、` ... → level 3
   - `(1)` / `（1）` / `(2)` / `（2）` ... → level 4
   - `①` / `②` ... → level 5

2. 构建标题树（缩进层级），验证父子关系是否连续。

3. 输出映射表 `{docx_dir}/{safe_name}-typ/heading_map.json`：
   ```json
   {
     "level_scheme": "chapter_section",  // or "flat"
     "mapping": {
       "第X章": 0,
       "一、": 1,
       "（一）": 2,
       "1.": 3,
       "(1)": 4,
       "①": 5
     },
     "headings": [
       {"text": "土壤属性图编制概述", "pattern": "一、", "level": 1, "page": 1},
       ...
     ]
   }
   ```

4. 将此 JSON 路径返回。
```

This mapping feeds into Step 1 (converter can emit correct heading levels) and Step 3 (bare-headings uses it as ground truth).

---

### Step 1: DOCX → TYP (C# converter)

Run in the main context via Bash. This step just prints progress — no context pressure.

```bash
cd ~/.claude/skills/docx2typ/scripts/docx2typ
dotnet run -- "<input.docx>"
```

Creates `{filename}-typ/` next to the input file:

```
{filename}-typ/
├── main.typ            # Format rules + content
├── .vscode/settings.json
├── figures/            # Extracted images
└── ref.bib             # Bibliography stub
```

---

### Step 2: Fix hardcoded 图N/表N → @label (ref2label)

**Spawn as a subagent** with `subagent_type: "general-purpose"`. This step edits `main.typ` in-place — the file diff can be large, but only the summary comes back.

**Subagent prompt** (replace `{main_typ}`, `{prefix}`):

```
Run the ref2label fix_refs.py script on a Typst file to convert hardcoded 图N/表N references to @label syntax.

Command:
  cd ~/.claude/skills/ref2label/scripts
  PYTHONIOENCODING=utf-8 python fix_refs.py --prefix={prefix} "{main_typ}"

Report:
- Number of table and image labels assigned
- Number of hardcoded references replaced
- Any warnings (unmatched references)
- Do NOT edit the file yourself — the script already does that.
```

If no chapter prefix is available, skip this step (ref2label needs unique per-chapter labels).

---

### Step 3: Scan bare headings + LLM review

**Spawn as a subagent** with `subagent_type: "general-purpose"`. The subagent runs `find_candidates.py`, then reviews each high-confidence candidate against surrounding context.

**Subagent prompt** (replace `{main_typ}`):

```
Run the bare-headings pipeline on a Typst file:
1. Run find_candidates.py to get candidate bare headings:
     cd ~/.claude/skills/bare-headings/scripts
     PYTHONIOENCODING=utf-8 python find_candidates.py "{main_typ}" > /tmp/candidates.json
2. Read the candidates JSON. For each HIGH-confidence candidate:
   - Read ~5 lines of context around the candidate line in the .typ file
   - Decide: is this really a heading, or just a list item / paragraph label / data annotation?
   - If it's a heading, what level should it be? (look at surrounding heading stack)
3. Apply edits: add =/==/=== markers to confirmed headings using the Edit tool.

Report:
- Total candidates found (high/medium/low)
- For each high-confidence candidate: verdict (heading/not heading) + action taken
- If level assignment was ambiguous, note it.
```

---

### Step 4: Insert images into table cells (new)

The C# converter extracts images to `figures/` but does **not** place them into table cells. DOCX tables with image grids (2‑col/3‑col composites of sub‑plots, common in Chinese reports) lose all images.

**Run in the main context** — write a Python script, or execute inline:

```bash
cd ~/.claude/skills/docx2typ/scripts
PYTHONIOENCODING=utf-8 python embed_table_images.py \
  --docx "{input.docx}" \
  --typ "{main_typ}" \
  --figures "{filename}-typ/figures"
```

Logic (implement in `embed_table_images.py`):
1. Parse DOCX XML — walk `<w:tbl>` elements, map each `a:blip` to its cell position (table# → row → col).
2. For each `#figure(kind: "table")` in the Typst file containing 2‑column image grids: load the cell‑text → image mapping from step 1.
3. Replace flat `[text],` cells with `[\n  #image("figures/figureN.png", width: 100%)\n  text\n],`.
4. For mixed‑content tables (e.g., 3‑col: name + env_vars + image): insert image cells into the correct column position.

**Key insight**: Image order in DOCX follows reading order (left‑to‑right, top‑to‑bottom), which matches Typst's flat `table()` cell order. Simple sequential mapping works: cell[0] → img#(start+0), cell[1] → img#(start+1), etc.

Report number of images inserted and any unmapped images.

---

### Step 5: Final cleanup — captions, references, escaping (new)

Run a Python script to fix remaining converter artifacts:

```bash
cd ~/.claude/skills/docx2typ/scripts
PYTHONIOENCODING=utf-8 python cleanup_typ.py "{main_typ}"
```

**Operations:**

| Issue | Fix |
|-------|-----|
| Hardcoded caption numbering `caption: [-1 xxx]` | Strip `-\d+` prefix → `caption: [xxx]` |
| ref2label suffix `@t-3-1-1` | Strip trailing `-\d+` → `@t-3-1` |
| ref2label + CJK `@t-3-1所示` | Add space → `@t-3-1 所示` |
| `~` as range indicator | Replace `~` → `\~` in content blocks |
| Long model‑param strings | Insert U+200B zero‑width space after commas to enable line‑breaking |

See step-by-step below if manual review is preferred.

---

### Alternative: one-shot wrapper (outside Claude)

For quick runs outside the Claude agent loop, the `convert.py` script runs all 3 steps sequentially:

```bash
cd ~/.claude/skills/docx2typ/scripts
python convert.py <input.docx>                    # auto-detect chapter prefix
python convert.py --prefix=07 <input.docx>        # explicit prefix
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

The converter outputs `=` / `==` / `===` / `====` / `=====` heading levels. Numbering is set via a custom function:

| Level | Typst | StyleId | Text pattern | Example |
|-------|-------|---------|-------------|---------|
| 1 | `=` | `1` or custom | `一、` | 一、土壤类型 |
| 2 | `==` | `2` or custom | `（一）` | （一）分类体系 |
| 3 | `===` | `4` or custom | `1.` / `1、` | 1. 土纲 |
| 4 | `====` | — | `(1)` / `（1）` | (1) 亚纲 |
| 5 | `=====` | — | `①` / `②` | ① 特征指标 |

Detection priority: (1) styleId (2) text pattern (3) font size hint. Custom styles (`af0`) no longer block heading detection.

### Why `#table()` with flat cells?

Typst's `#table()` uses a flat list of content blocks (one per cell), unlike LaTeX's row-based `&` / `\\` syntax. This maps naturally to the DOCX traversal order. Merged cells use `#table.cell(colspan: N, rowspan: M)[...]`. The `stroke: 0.5pt` option provides Word-style grid borders.

### Why `fr` column widths?

Column weight ratios are derived from the DOCX `tblGrid` widths and expressed as `(3fr, 5fr, 2fr, ...)`. Typst's `fr` unit divides the available space proportionally — simpler and more robust than LaTeX's `X` column weight system.

### Why no `.latexmkrc` or format file?

Typst compiles with a single `typst compile` command. The Tinymist VSCode extension handles preview and auto-compilation. No build automation scripts needed.

## Section Structure Detection

Two detection paths run in parallel (identical to docx2tex):

### Path A: style-priority (styleId)

| Style ID | Mapping | Text fallback | Example |
|----------|---------|--------------|---------|
| `1` | `= ` | `一、` | 一、工作背景 |
| `2` | `== ` | `（一）` | (一) 工作组织 |
| `4` | `=== ` | `1.` / `1、` | 1. 培肥改良土壤 |
| — | `==== ` | `(1)` / `（1）` | (1) 技术措施 |
| — | `===== ` | `①` / `②` | ① 石灰施用 |

### Path B: text fallback (no style)

Text pattern fallback runs regardless of styleId (including `af0` custom styles):

| Text pattern | Level | Output |
|-------------|-------|--------|
| `一、` / `二、` ... | 1 | `= ` |
| `（一）` / `（二）` ... | 2 | `== ` |
| `1.` / `1．` / `1、` (1-2 digit only) | 3 | `=== ` |
| `(1)` / `（1）` ... | 4 | `==== ` |
| `①` / `②` ... | 5 | `===== `

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

Typst content blocks `[...]` require escaping of `\`, `#`, `[`, `]`, `<`, `>`, `~` — `<>` is interpreted as a label delimiter; `~` is a non-breaking space in content mode. Chinese characters need no escaping at all.

| Character | Escape | Reason |
|-----------|--------|--------|
| `\` | `\\` | Escape character |
| `#` | `\#` | Code expression prefix |
| `[` | `\[` | Content block delimiter |
| `]` | `\]` | Content block delimiter |
| `<` | `\<` | Label delimiter |
| `>` | `\>` | Label delimiter |
| `~` | `\~` | Non‑breaking space in content mode |
| `_` | `\_` | Italic marker in content mode |

### Critical Typst Gotchas (from real-world testing)

1. **`table.cell()` has NO `#` prefix** — inside `table()` arguments, code mode is active; `table.cell(colspan: 2)[...]` is correct, `#table.cell(...)` fails
2. **`#figure()` blocks page breaking by default** — must add `#show figure.where(kind: "table"): set block(breakable: true)` for tables
3. **`#figure()` requires `supplement`** in Typst ≥0.14 — set `supplement: [表]` / `supplement: [图]` per call
4. **`{ set text(); table() }` code blocks prevent page breaking** — use `text(size: Npt, table(...))` wrapper instead
5. **Trailing commas inside code blocks are illegal** — `{ set text(); table(...), }` fails; comma belongs after block close `},` in the parent `#figure()` call
6. **Zero‑width spaces for long strings** — Typst wraps text at spaces and hyphens, but code‑like parameter strings (`alpha=0.5,l1_ratio=0.6,...`) have no natural break points. Insert U+200B (zero‑width space) after commas to enable line‑breaking. The `cleanup_typ.py` script in Step 5 handles this.
7. **WPS 垂直合并全部标记 restart** — WPS/DOCX 在垂直合并中把所有 `<w:vMerge>` 写为 `w:val="restart"`（标准做法是第一个 `restart`、后续 `continue`）。导致转换器按 OpenXML API 读不到 continue 单元格，多算 rowspan。实际表现：rowspan 值比实际多 1 行，最后一行出现冲突的单元格。**修复**：在输出循环中验证 `pos + colspan <= colCount`，若超限则 clamp colspan 或整表拍平。
8. **`_` 下划线未转义** — Typst content 模式下 `_` 是斜体标记。`FormatCellTypst` 需额外 `.Replace("_", "\\_")`，否则包含 `_` 的单元格文本（如 `TPI_标准`）会触发 unclosed delimiter 错误。
9. **`counter(figure.where(kind: image))` vs `kind: "image"`** — 图/表编号计数器重置时，`kind` 必须与 `#figure()` 调用中用的完全相同。`#figure(kind: "image")` 用的是字符串，`counter(figure.where(kind: "image"))` 也必须用字符串。若写成 `kind: image`（无引号）则解析为函数，类型不匹配，计数器永不重置。

## Caption Handling

### Table captions
Strips `表N.N` or `附表N` prefixes. Also strips hardcoded `-N` numbering (e.g., `caption: [-1 土壤属性表]` → `[土壤属性表]`) — Typst auto‑numbering handles the rest. Wrapped in `#figure(table(...), caption: [...], kind: "table")`.

### Figure captions
Strips `图N(.N)(a-z)?` prefixes and hardcoded `-N` prefixes. Wrapped in `#figure(image(...), caption: [...], kind: "image")`.

### Post‑conversion cleanup
Run Step 5 `cleanup_typ.py` to catch any remaining `-N` prefixes that the converter missed.

## Image Handling

- **XML direct scan**: extracts images by scanning paragraph XML for `a:blip` references
- **TIFF auto-conversion**: TIFF images are automatically converted to PNG using `System.Drawing.Common` during extraction (Typst does not support TIFF natively)
- Images extracted to `figures/` as `figure1.jpg`, etc.
- Typst natively supports SVG, PNG, JPG

## Table Page Breaking

Typst `table()` breaks across pages automatically — no special configuration needed. Long tables with hundreds of rows will page-break naturally.

However, if a table is wrapped inside a `{ }` code block (e.g., for font-size scoping), the block prevents page breaking. The converter uses `text(size: Npt, table(...))` instead, which preserves both the font size and page-breaking behavior.

## Known Limitations

The post-processing pipeline (see above) auto-fixes some known issues (hardcoded `图N`/`表N` cross-references via ref2label). The following still need manual attention:

- **Bare headings**: the subagent workflow (Step 3) reviews and fixes high-confidence candidates automatically. Low/medium confidence candidates and edge cases may still need manual review. Use Pre-step B's `heading_map.json` as ground truth to verify.
- **Numbering style**: hardcoded as `#set heading(numbering: "一、")`. Adjust manually for other styles. For multi-chapter reports, shift levels: `第X章` → `=`, `一、` → `==`, `（一）` → `===`, `1.` → `====`, `(1)` → `=====`.
- **Image-grid tables**: the converter extracts images from DOCX table cells but does NOT place them back. Step 4 (`embed_table_images.py`) handles this, but may miss tables with irregular merges. Manual verification recommended for composite figure grids.
- **Heading mapping drift**: DOCX style IDs are often unreliable (custom `af0`, missing IDs). Don't trust them exclusively — always cross-check with Pre-step B's semantic heading extraction.
- **Cell background colors**: not yet mapped from `tcPr/shd` to Typst's `fill:` option.
- **SVG images**: extracted as `.bin` — Typst supports SVG, but the DOCX's SVG content type is not yet detected. Manually rename to `.svg` if needed.
- **`.doc` (old binary)**: not supported. Convert to `.docx` first.

## Requirements

- .NET SDK 8.0+
- NuGet: `DocumentFormat.OpenXml` 3.2.0 (auto-restored)
- Typst CLI (`typst compile`) or Tinymist VSCode extension
