#!/usr/bin/env python3
"""
docx2typ full pipeline: C# converter → ref2label → bare-headings scan.

Usage: python convert.py [--prefix=XX] <input.docx>

Steps:
1. Run the C# docx2typ converter (DOCX → TYP)
2. Fix hardcoded 图N/表N → @label (ref2label, in-place)
3. Scan for bare headings (bare-headings, report only)
"""

import sys
import os
import re
import subprocess
import json
import shutil

SKILLS_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOCX2TYP_DIR = os.path.join(SKILLS_DIR, "docx2typ", "scripts", "docx2typ")
FIX_REFS = os.path.join(SKILLS_DIR, "ref2label", "scripts", "fix_refs.py")
FIND_CANDIDATES = os.path.join(SKILLS_DIR, "bare-headings", "scripts", "find_candidates.py")

CN_NUMS = {
    '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
    '六': '6', '七': '7', '八': '8', '九': '9', '十': '10',
    '零': '0', '〇': '0',
}


def detect_prefix(filename):
    """Auto-detect chapter prefix from filename patterns."""
    m = re.search(r'第([一二三四五六七八九十\d]+)章', filename)
    if m:
        cn = m.group(1)
        if cn.isdigit():
            return cn.zfill(2)
        return CN_NUMS.get(cn)
    m = re.search(r'ch(\d+)', filename, re.IGNORECASE)
    if m:
        return m.group(1).zfill(2)
    return None


def main():
    args = sys.argv[1:]
    prefix = None
    docx_path = None
    for a in args:
        if a.startswith('--prefix='):
            prefix = a.split('=', 1)[1]
        elif not a.startswith('-'):
            docx_path = a

    if not docx_path:
        print("Usage: python convert.py [--prefix=XX] <input.docx>", file=sys.stderr)
        sys.exit(1)

    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        print(f"ERROR: file not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    # Auto-detect prefix from filename
    if not prefix:
        prefix = detect_prefix(os.path.basename(docx_path))

    # ── Step 1: C# converter ──
    print("=" * 60)
    print("Step 1/3: DOCX → TYP (C# converter)")
    print("=" * 60)
    result = subprocess.run(
        ["dotnet", "run", "--", docx_path],
        cwd=DOCX2TYP_DIR,
    )
    if result.returncode != 0:
        print("\nERROR: C# converter failed", file=sys.stderr)
        sys.exit(1)

    # Locate output (same logic as Program.cs)
    file_name = os.path.splitext(os.path.basename(docx_path))[0]
    safe_name = re.sub(r'[^\w\- ]', '_', file_name)
    output_dir = os.path.join(os.path.dirname(docx_path), f"{safe_name}-typ")
    main_typ = os.path.join(output_dir, "main.typ")

    if not os.path.exists(main_typ):
        print(f"\nERROR: output not found at {main_typ}", file=sys.stderr)
        sys.exit(1)

    # ── Step 2: ref2label ──
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'

    if prefix:
        print(f"\n{'='*60}")
        print(f"Step 2/3: Fix hardcoded 图/表 refs (ref2label, prefix={prefix})")
        print("=" * 60)
        result = subprocess.run(
            ["python", FIX_REFS, f"--prefix={prefix}", main_typ],
            env=env,
        )
        if result.returncode != 0:
            print("WARNING: ref2label had errors (file may still be usable)")
    else:
        print(f"\n{'='*60}")
        print("Step 2/3: Fix hardcoded 图/表 refs — SKIPPED")
        print("=" * 60)
        print("WARNING: Could not detect chapter prefix from filename.")
        print("  Use --prefix=XX to specify manually (e.g. --prefix=07).")
        print("  Filename should contain 第X章 or ch0X pattern for auto-detection.")

    # ── Step 3: bare-headings scan ──
    print(f"\n{'='*60}")
    print("Step 3/3: Scan for bare headings (bare-headings)")
    print("=" * 60)
    result = subprocess.run(
        ["python", FIND_CANDIDATES, main_typ],
        capture_output=True, text=True, env=env,
    )
    try:
        candidates = json.loads(result.stdout)
    except json.JSONDecodeError:
        print("WARNING: Could not parse bare-headings output")
        candidates = []

    high = [c for c in candidates if c['confidence'] == 'high']
    med = [c for c in candidates if c['confidence'] == 'medium']
    low_count = len(candidates) - len(high) - len(med)
    print(f"Found {len(candidates)} candidates: {len(high)} high, {len(med)} medium, {low_count} low")

    if high:
        print("\n⚠  High-confidence bare headings (run bare-headings skill to fix):")
        for c in high:
            print(f"  L{c['line']:4d}: {c['text'][:70]}")

    # ── Summary ──
    print(f"\n{'='*60}")
    print("Conversion complete!")
    print(f"  Output: {main_typ}")
    if prefix:
        print(f"  ref2label: applied (prefix={prefix})")
    else:
        print(f"  ref2label: skipped (no prefix)")
    if high:
        print(f"  bare-headings: {len(high)} candidates to review")
    else:
        print(f"  bare-headings: clean")
    print("=" * 60)


if __name__ == '__main__':
    main()
