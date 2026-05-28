#!/usr/bin/env python3
"""
Post-conversion cleanup for DOCX-to-Typst output.

Fixes known converter artifacts:
1. Hardcoded caption numbering: caption: [-N xxx] → caption: [xxx]
2. ref2label suffix mismatch: @t-3-1-1 → @t-3-1
3. ref2label + CJK: @t-3-1所示 → @t-3-1 所示
4. Tilde escaping: ~ → \~ in content blocks
5. ZWS insertion: zero-width spaces after commas in long strings
6. Duplicate figure caption: caption: [] → caption filled from context

Usage:
  python cleanup_typ.py <main.typ> [--figures-dir=<path>]

Output: modifies main.typ in-place.
"""

import sys
import os
import re


def parse_args():
    args = sys.argv[1:]
    typ_path = None
    figures_dir = None
    for a in args:
        if a.startswith('--figures-dir='):
            figures_dir = a.split('=', 1)[1]
        elif not a.startswith('-'):
            typ_path = a

    if not typ_path:
        print("Usage: python cleanup_typ.py <main.typ> [--figures-dir=<path>]", file=sys.stderr)
        sys.exit(1)
    typ_path = os.path.abspath(typ_path)
    if not figures_dir:
        figures_dir = os.path.join(os.path.dirname(typ_path), 'figures')
    if not os.path.exists(typ_path):
        print(f"ERROR: {typ_path} not found", file=sys.stderr)
        sys.exit(1)
    return typ_path, figures_dir


def fix_hardcoded_caption_numbering(content: str) -> tuple[str, int]:
    """Remove -N prefixes from caption fields."""
    pattern = r'(caption:\s*\[)-\d+\s*'
    matches = list(re.finditer(pattern, content))
    content = re.sub(pattern, r'\1', content)
    return content, len(matches)


def fix_ref_suffixes(content: str) -> tuple[str, int]:
    """Remove trailing -N from @ref references: @t-3-1-1 → @t-3-1."""
    pattern = r'@([tf]-3-\d+)-\d+'
    matches = list(re.finditer(pattern, content))
    content = re.sub(pattern, r'@\1', content)
    return content, len(matches)


def fix_ref_cjk_spacing(content: str) -> tuple[str, int]:
    """Add space between @ref and following CJK character."""
    pattern = r'(@[tf]-3-\d+)([一-鿿])'
    matches = list(re.finditer(pattern, content))
    content = re.sub(pattern, r'\1 \2', content)
    return content, len(matches)


def fix_tilde_escaping(content: str) -> tuple[str, int]:
    """
    Escape ~ → \\~ in Typst content blocks.
    Only escapes tildes that are used as range indicators (between digits or Chinese).
    """
    # Match ~ between digits (like 2~0.2mm), digits+Chinese, or Chinese+digits
    # Use a lookahead/lookbehind approach
    count = 0
    result = []
    i = 0
    while i < len(content):
        if content[i] == '~':
            # Only escape tildes that are in content mode (inside [...] blocks)
            # Check context: are we inside a [...] block?
            # Simple heuristic: if preceded/followed by digit or CJK, it's a range indicator
            prev = content[i - 1] if i > 0 else ''
            next_c = content[i + 1] if i + 1 < len(content) else ''
            is_range = (
                prev.isdigit() or ('一' <= prev <= '鿿') or
                next_c.isdigit() or ('一' <= next_c <= '鿿')
            )
            if is_range and content[i - 1] != '\\':
                result.append('\\~')
                count += 1
            else:
                result.append('~')
        else:
            result.append(content[i])
        i += 1
    return ''.join(result), count


def fix_zws_long_strings(content: str) -> tuple[str, int]:
    """
    Insert zero-width spaces (U+200B) after commas in long parameter strings
    inside table cells. Targets cells with model parameters (contain = and ,).
    """
    # Match [content] cells that have 2+ = signs (model parameter pattern)
    count = 0

    def add_zws(m):
        inner = m.group(1)
        if inner.count('=') >= 2 and len(inner) > 80:
            inner = inner.replace(',', ',​')
            count_global = m.group(0).count('=')
            nonlocal count
            count += inner.count('​')
        return '[' + inner + ']'

    # Use nonlocal
    ns = {'count': 0}

    def replacer(m):
        inner = m.group(1)
        if inner.count('=') >= 2 and len(inner) > 80:
            inner = inner.replace(',', ',​')
            ns['count'] += inner.count('​')
        return '[' + inner + ']'

    content = re.sub(r'\[([^\[\]]+)\]', replacer, content)
    return content, ns['count']


def fix_empty_caption(content: str) -> tuple[str, int]:
    """Replace empty caption: [] with a placeholder if it's the first figure (technical route)."""
    if 'caption: []' in content:
        content = content.replace('caption: []', 'caption: [图]', 1)
        return content, 1
    return content, 0


def main():
    typ_path, figures_dir = parse_args()

    with open(typ_path, 'r', encoding='utf-8') as f:
        content = f.read()

    report = []

    # 1. Hardcoded caption numbering
    content, n = fix_hardcoded_caption_numbering(content)
    report.append(f"Caption -N prefixes removed: {n}")

    # 2. ref2label suffix mismatch
    content, n = fix_ref_suffixes(content)
    report.append(f"ref2label suffixes stripped: {n}")

    # 3. ref2label + CJK spacing
    content, n = fix_ref_cjk_spacing(content)
    report.append(f"ref2label + CJK spacing fixed: {n}")

    # 4. Tilde escaping
    content, n = fix_tilde_escaping(content)
    report.append(f"Tildes escaped: {n}")

    # 5. ZWS for long strings
    content, n = fix_zws_long_strings(content)
    report.append(f"ZWS inserted in long strings: {n}")

    # 6. Empty caption
    content, n = fix_empty_caption(content)
    if n:
        report.append(f"Empty captions filled: {n}")

    with open(typ_path, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Cleanup complete:")
    for r in report:
        print(f"  • {r}")


if __name__ == '__main__':
    main()
