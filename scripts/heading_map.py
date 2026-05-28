#!/usr/bin/env python3
"""
Semantic heading extraction from DOCX — text-pattern only.

Reads a .docx file, extracts all paragraphs whose text matches Chinese
heading numbering patterns, builds a hierarchical tree, and outputs a
JSON mapping for the docx2typ pipeline.

Heading patterns (reliable for Chinese government/academic reports):
  Level 0: 第X章 / 第X节 / 第X部分
  Level 1: 一、 二、 三、 ...
  Level 2: （一）（二）（三）... or (一)(二)(三)...
  Level 3: 1. 1、 N. N、 (short, with whitespace after delimiter)
  Level 4: (1) (2) （1）（2）... (short)
  Level 5: ① ② ③ ...

StyleId and Word auto-numbering (numPr) are NOT used — they are
unreliable in WPS-authored documents. Use the separate bare-headings
pipeline (Step 3) for style-based heading detection.

Usage:
  python heading_map.py <input.docx> [-o output.json]
"""

import sys
import os
import re
import json
import zipfile
from xml.etree import ElementTree as ET

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

CN_NUMS = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
           '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}


def extract_paragraphs(docx_path):
    """Extract paragraph text from DOCX in document order."""
    with zipfile.ZipFile(docx_path) as z:
        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
    body = tree.getroot().find('.//w:body', NS)
    paragraphs = []
    for para in body.findall('.//w:p', NS):
        texts = []
        for t in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
            if t.text:
                texts.append(t.text)
        text = ''.join(texts).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def is_heading_title(title):
    """Filter out false positives: data values, measurements, running text."""
    if not title or len(title) > 60:
        return False
    # Reject model parameters or measurement values
    if '=' in title:
        return False
    if re.search(r'~\d+\.?\d*\s*mm', title):
        return False
    if re.search(r'^\d+[\.\d]*\s*(cm|mm|亩|g/kg|mg/kg)', title):
        return False
    # Reject running text (3+ common particles in a title-length string)
    markers = sum(1 for m in '的了是在有为与或并和' if m in title)
    if markers >= 3 and len(title) > 15:
        return False
    return True


def detect_heading(text):
    """
    Returns (level, pattern, clean_title) or (None, '', text).
    """
    # Level 0: 第X章 / 第X节
    m = re.match(r'^第([一二三四五六七八九十\d]+)([章节部篇])', text)
    if m:
        title = text[m.end():].lstrip()
        return (0, 'chapter', title or text)

    # Level 1: 一、二、三、...
    m = re.match(r'^([一二三四五六七八九十]+)[、，,]\s*', text)
    if m and is_heading_title(text[m.end():]):
        return (1, 'section', text[m.end():])

    # Level 2: （一）（二）... or (一)(二)...
    m = re.match(r'^[（(]([一二三四五六七八九十]+)[）)]\s*', text)
    if m and is_heading_title(text[m.end():]):
        return (2, 'subsection', text[m.end():])

    # Level 3: 1. 1、 (with whitespace after delimiter, short title)
    m = re.match(r'^(\d{1,2})[.、．]\s+', text)
    if m:
        title = text[m.end():]
        if len(title) <= 40 and is_heading_title(title):
            return (3, 'sub_subsection', title)

    # Level 4: (1) (2) （1）（2） (short title)
    m = re.match(r'^[（(](\d{1,2})[）)]\s*', text)
    if m:
        title = text[m.end():]
        if len(title) <= 30 and is_heading_title(title):
            return (4, 'paragraph', title)

    # Level 5: ①②③...
    m = re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩]', text)
    if m:
        title = text[m.end():].lstrip()
        if len(title) <= 35 and is_heading_title(title):
            return (5, 'sub_paragraph', title)

    return (None, '', text)


def build_tree(headings):
    """Build hierarchical tree, inserting placeholders for skipped levels."""
    tree = []
    stack = []

    for h in headings:
        lv = h['level']
        # Clamp excessive skip: if jump > 1, reset to prev+1
        while stack and stack[-1][0] >= lv:
            stack.pop()

        node = {'text': h['title'], 'level': lv, 'pattern': h['pattern']}

        if stack:
            parent = stack[-1][1]
            parent.setdefault('children', []).append(node)
        else:
            tree.append(node)

        stack.append((lv, node))

    return tree


def main():
    args = sys.argv[1:]
    docx_path = None
    output_path = None
    for a in args:
        if not a.startswith('-'):
            if docx_path is None:
                docx_path = a
            else:
                output_path = a

    if not docx_path:
        print("Usage: python heading_map.py <input.docx> [output.json]", file=sys.stderr)
        sys.exit(1)

    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        print(f"ERROR: file not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    # Default output path
    if output_path is None:
        base = os.path.splitext(docx_path)[0]
        out_dir = os.path.join(os.path.dirname(docx_path),
                               re.sub(r'[^\w\- ]', '_', os.path.basename(base)) + '-typ')
        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, 'heading_map.json')

    paragraphs = extract_paragraphs(docx_path)

    # Detect headings
    headings = []
    for text in paragraphs:
        level, pattern, title = detect_heading(text)
        if level is not None:
            headings.append({'level': level, 'pattern': pattern,
                             'title': title or text, 'raw': text})

    # Level scheme
    has_chapter = any(h['level'] == 0 for h in headings)
    levels_found = sorted(set(h['level'] for h in headings))

    # Typst mapping
    level_labels = {0: '=   # 第X章', 1: '==  # 一、', 2: '===  # （一）',
                    3: '==== # 1.',   4: '===== # (1)',  5: '====== # ①'}

    tree = build_tree(headings)

    result = {
        'level_scheme': 'chapter_section' if has_chapter else 'flat',
        'detected_levels': levels_found,
        'typst_mapping': {str(k): level_labels.get(k, '') for k in levels_found},
        'heading_count': len(headings),
        'headings': headings,
        'tree': tree,
    }

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Heading map written to {output_path}")
    print(f"  Headings detected: {len(headings)}")
    print(f"  Levels: {levels_found}")
    print(f"  Scheme: {result['level_scheme']}")
    print(f"\nTypst level mapping:")
    for lv in levels_found:
        print(f"  Level {lv}: {level_labels.get(lv, '')}")

    if len(headings) < 3:
        print("\nNOTE: Few headings detected by text pattern. This DOCX likely uses")
        print("  Word auto-numbering or style-based headings. Run the bare-headings")
        print("  pipeline (Step 3) to detect headings from Typst context.")


if __name__ == '__main__':
    main()
