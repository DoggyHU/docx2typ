#!/usr/bin/env python3
"""
Embed images into DOCX-converted Typst table cells.

The C# docx2typ converter extracts all images from a DOCX to the figures/
directory, but does NOT place them into table cells. Tables that contain
image grids (common in Chinese scientific reports: anomaly detection plots,
normal distribution comparisons, variogram fitting, etc.) lose all images.

This script:
1. Parses the DOCX XML to find which images belong in which table cells
2. Reads the Typst file produced by the converter
3. For each table with embedded images, inserts `#image()` calls into cells
4. For pure image-grid tables (2-col composites), wraps each cell with image + text
5. For mixed tables (e.g. 3-col: name + env_vars + image), inserts image column cells

Usage:
  python embed_table_images.py --docx input.docx --typ main.typ [--figures figures/]

Output: modifies main.typ in-place.
"""

import sys
import os
import re
import zipfile
from xml.etree import ElementTree as ET

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def parse_args():
    args = sys.argv[1:]
    docx = typ = figures = None
    for a in args:
        if a.startswith('--docx='):
            docx = a.split('=', 1)[1]
        elif a.startswith('--typ='):
            typ = a.split('=', 1)[1]
        elif a.startswith('--figures='):
            figures = a.split('=', 1)[1]

    if not docx or not typ:
        print("Usage: python embed_table_images.py --docx=<input.docx> --typ=<main.typ> [--figures=<figures/>]")
        sys.exit(1)

    # Auto-detect figures dir if not provided
    if not figures:
        figures = os.path.join(os.path.dirname(typ), 'figures')

    return docx, typ, figures


def build_image_table_map(docx_path):
    """
    Parse DOCX XML to map image positions to table cells.
    Returns: dict of table_index -> list of (image_number, row, text_content)
    """
    with zipfile.ZipFile(docx_path) as z:
        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
    root = tree.getroot()

    parent_map = {}
    for parent in root.iter():
        for child in parent:
            parent_map[child] = parent

    image_idx = 0
    table_num = 0
    row_num = 0
    result = {}

    for elem in root.iter():
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else ''
        if not tag:
            continue

        if tag == 'tbl':
            table_num += 1
            row_num = 0
            result.setdefault(table_num, [])

        if tag == 'tr':
            row_num += 1

        if tag == 'blip':
            image_idx += 1
            # Walk up to find containing cell
            p = elem
            texts = []
            while p is not None:
                pt = p.tag.split('}')[-1] if '}' in p.tag else ''
                if pt == 'tc':
                    for t in p.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                        if t.text:
                            texts.append(t.text)
                    break
                p = parent_map.get(p)

            text = ''.join(texts)[:120] if texts else ''
            result.setdefault(table_num, []).append((image_idx, row_num, text))

    return result, image_idx


def get_image_extension(figures_dir, n):
    """Check if figureN exists as .png or .jpg."""
    png = os.path.join(figures_dir, f"figure{n}.png")
    jpg = os.path.join(figures_dir, f"figure{n}.jpg")
    if os.path.exists(png):
        return '.png'
    if os.path.exists(jpg):
        return '.jpg'
    return '.png'  # default guess


def process_pure_image_grid(content, label, img_start, cell_texts_suffix, ext):
    """
    For a 2-column image grid table (e.g., t-3-1 anomaly detection).
    Each cell has: [text], → becomes [\n  #image("figures/figureN.ext")\n  text\n],
    """
    for i, text in enumerate(cell_texts_suffix):
        img_n = img_start + i
        ext_actual = ext or get_image_extension(
            os.path.join(os.path.dirname(sys.argv[2] if len(sys.argv) > 2 else ''), 'figures'),
            img_n
        )
        old = f'[{text}],'
        new = f'[\n  #image("figures/figure{img_n}{ext_actual}", width: 100%)\n  {text}\n],'
        if old in content:
            content = content.replace(old, new, 1)
    return content


def main():
    docx_path, typ_path, figures_dir = parse_args()

    print(f"Parsing DOCX: {docx_path}")
    table_images, total_images = build_image_table_map(docx_path)

    print(f"  Found {total_images} images in {len(table_images)} tables")

    # Summarize tables that have images
    print(f"\nTables with images:")
    for tbl, imgs in sorted(table_images.items()):
        if imgs:
            print(f"  Table {tbl}: {len(imgs)} images (img#{imgs[0][0]}–#{imgs[-1][0]})")

    with open(typ_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # ── Table 1: anomaly detection (2-col image grid) ──
    # DOCX images start at img#2. Cells match COMMON_TEXTS + "空间异常点检测图"
    common = [
        '2~0.2mm颗粒含量', '0.2~0.02mm颗粒含量',
        '0.02~0.002mm颗粒含量', '0.002mm以下颗粒含量',
        '耕作层厚度', '土壤容重平均值', 'pH', '阳离子交换量',
        '交换性盐基总量', '水溶性盐总量', '有机质', '全氮', '全磷', '全钾',
        '有效磷', '速效钾', '缓效钾', '有效硫', '有效铁', '有效锰',
        '有效铜', '有效锌', '有效硼', '有效钼',
        '交换性钙', '交换性镁', '有效硅', '全硒',
    ]

    grid_tables = [
        # (label, img_start, suffix, ext)
        ('t-3-1', 2, '空间异常点检测图', '.png'),
        ('t-3-5', 30, '正态分布转换前后对比图', '.png'),
        ('t-3-13', 86, '变异函数拟合图', '.png'),
    ]

    for label, img_start, suffix, ext in grid_tables:
        texts = [c + suffix for c in common]
        count_before = content.count('#image(')
        for i, text in enumerate(texts):
            img_n = img_start + i
            old = f'[{text}],'
            new = f'[\n  #image("figures/figure{img_n}{ext}", width: 100%)\n  {text}\n],'
            if old in content:
                content = content.replace(old, new, 1)
        count_after = content.count('#image(')
        inserted = count_after - count_before
        print(f"  {label}: {inserted} images inserted")

    # ── Mixed tables: data tables where col 3 is an image (t-3-12: 入模环境变量) ──
    # Detect by finding #figure blocks with columns:(3fr, 4fr, 1fr) and header containing 贡献度
    # These need: [name], [env], [image] per row
    t12_pattern = re.compile(
        r'(#figure\(\n  text\(size: 10\.5pt,\n    table\(\n      columns: \(3fr, 4fr, 1fr\),\n      stroke: 0\.5pt,\n      table\.header\(.*?\),\n)(.*?)(\n    \)\n  \),\n)', re.DOTALL
    )
    m = t12_pattern.search(content)
    if m:
        # Find all [content], cells in the block
        cells_raw = re.findall(r'\[(.*?)\],', m.group(2))
        if len(cells_raw) >= 56:  # 28 name + 28 env at minimum
            # Rebuild with images
            names = cells_raw[0::2]
            envs = cells_raw[1::2]
            new_cells = []
            for i in range(min(len(names), len(envs), 28)):
                img_n = 58 + i
                new_cells.append(f'[{names[i]}]')
                new_cells.append(f'[{envs[i]}]')
                new_cells.append(f'[#image("figures/figure{img_n}.jpg", width: 100%)]')
            new_block = m.group(1) + '\n      ' + \
                        ',\n      '.join(new_cells) + ',' + \
                        m.group(3)
            content = content.replace(m.group(0), new_block, 1)
            print(f"  t-3-12 (mixed): {len(names[:28])} rows with images 58-85")

    # ── T17: internal validation (3 images per cell) ──
    t17_texts = [
        'pH（样点、土地利用、土壤类型、地形地貌）',
        '有机质（样点、土地利用、土壤类型、地形地貌）',
        '全氮（样点、土地利用、土壤类型、地形地貌）',
        '速效钾（样点、土地利用、土壤类型、地形地貌）',
        '有效磷（样点、土地利用、土壤类型、地形地貌）',
    ]
    for j, text in enumerate(t17_texts):
        a, b, c = 114 + j*3, 115 + j*3, 116 + j*3
        old = f'table.cell(colspan: 3)[{text}]'
        new = (f'table.cell(colspan: 3)[\n'
               f'  #image("figures/figure{a}.png", width: 30%)\n'
               f'  #image("figures/figure{b}.png", width: 30%)\n'
               f'  #image("figures/figure{c}.png", width: 30%)\n'
               f'  {text}\n]')
        if old in content:
            content = content.replace(old, new, 1)
            print(f"  t-3-17 (internal val): row {j+1} with images {a}-{c}")

    with open(typ_path, 'w', encoding='utf-8') as f:
        f.write(content)

    total = content.count('#image(')
    print(f"\nDone. Total #image() calls in .typ file: {total}")


if __name__ == '__main__':
    main()
