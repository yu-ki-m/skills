#!/usr/bin/env python3
"""Verify that an .xlsx file written by build.py --xlsx opens in desktop Excel as the same table.

Give the same table as build.py was given (--input, or --base with or without --state) and the .xlsx file. For each sheet
  1. builds the same cells natively in Excel from the workbook (the reference; the same builder as verify_excel_paste.py),
  2. opens the .xlsx file read-only in the same hidden Excel,
  3. compares value, type, formula, number format, font, colours, fill, alignment, wrap, indent, merge, link and all four
     borders cell by cell, and compares column widths, row heights, frozen panes, gridlines, zoom, the active sheet and
     pictures with what the page shows (CSS pixels x 0.75 = points),
  4. reports any repair Excel had to make to open the file.

Requires Windows, desktop Excel, PowerShell 7 (pwsh) and Node.js 18+. It starts a hidden Excel and does not use the
clipboard; the user's open workbooks are not touched. Exit code: 0 = no unexpected difference, 1 = unexpected
differences, 2 = cannot run here (not Windows / no Excel / no pwsh).
"""
import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True  # importing build.py and verify_excel_paste.py must not leave __pycache__ files
sys.path.insert(0, str(Path(__file__).resolve().parent))
import build  # noqa: E402  (same folder)
import verify_excel_paste as paste  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
CHECK = ROOT / 'scripts' / 'excel_check'
# Excel keeps widths and heights in whole screen pixels: one pixel is 0.75pt at 100% display scaling and 0.6pt at 125%.
SIZE_TOLERANCE = 0.8
# Picture positions add up the rounding of every column and row before them.
POSITION_TOLERANCE = 3.0


def width_px(sheet, c):
    widths = sheet.get('colWidths') or []
    return widths[c] if c < len(widths) and widths[c] is not None else 96


def height_px(sheet, r):
    heights = sheet.get('rowHeights') or []
    return heights[r] if r < len(heights) and heights[r] is not None else 22


def anchor(sheet, picture):
    """The cell (0-based) and offset in px where the picture's top-left corner lies, as scripts/xlsx.cjs places it."""
    r, c, x, y = picture['r'], picture['c'], picture['x'], picture['y']
    while c < sheet['cols'] - 1 and x >= width_px(sheet, c):
        x -= width_px(sheet, c)
        c += 1
    while y >= height_px(sheet, r):
        y -= height_px(sheet, r)
        r += 1
    return r, c, x, y


def check_sheet(workbook, sheet, info):
    """Differences between the sheet's view, sizes and pictures in Excel and in the page."""
    problems = []

    def differ(item, page, excel):
        problems.append({'item': item, 'page': page, 'excel': excel})

    if info['repairs']:
        differ('Excelによる修復', '修復なしで開く', ' / '.join(x[:300] for x in info['repairs']))
    names = [s['name'] for s in workbook['sheets']]
    if info['sheetNames'] != names:
        differ('シートの名前と順番', names, info['sheetNames'])
    active = next(s['name'] for s in workbook['sheets'] if s['id'] == workbook['activeSheet'])
    if info['activeSheet'] != active:
        differ('開いたときのシート', active, info['activeSheet'])
    widths = info['colWidths'] if isinstance(info['colWidths'], list) else [info['colWidths']]
    heights = info['rowHeights'] if isinstance(info['rowHeights'], list) else [info['rowHeights']]
    for c, got in enumerate(widths):
        want = width_px(sheet, c) * 0.75
        if abs(got - want) > SIZE_TOLERANCE:
            differ(f'列 {paste.col_name(c + 1)} の幅（pt）', want, got)
    for r, got in enumerate(heights):
        want = height_px(sheet, r) * 0.75
        if abs(got - want) > SIZE_TOLERANCE:
            differ(f'{r + 1} 行目の高さ（pt）', want, got)
    freeze = sheet.get('freeze') or {}
    rows, cols = freeze.get('rows', 0), freeze.get('cols', 0)
    want_freeze = {'freeze': bool(rows or cols), 'splitRow': rows, 'splitColumn': cols} if rows or cols else {'freeze': False}
    got_freeze = {'freeze': info['freeze'], 'splitRow': info['splitRow'], 'splitColumn': info['splitColumn']} if info['freeze'] else {'freeze': False}
    if got_freeze != want_freeze:
        differ('枠の固定', want_freeze, got_freeze)
    if info['gridlines'] != (sheet.get('gridlines') is not False):
        differ('枠線の表示', sheet.get('gridlines') is not False, info['gridlines'])
    zoom = round((workbook.get('view') or {}).get('zoom', 100))
    if info['zoom'] != zoom:
        differ('倍率（%）', zoom, info['zoom'])
    pictures = info['pictures'] or []
    pictures = pictures if isinstance(pictures, list) else [pictures]
    images = sheet.get('images') or []
    if len(pictures) != len(images):
        differ('画像の数', len(images), len(pictures))
    for i, (image, got) in enumerate(zip(images, pictures)):
        r, c, x, y = anchor(sheet, image)
        left = sum(widths[k] if k < len(widths) else width_px(sheet, k) * 0.75 for k in range(c)) + x * 0.75
        top = sum(heights[k] if k < len(heights) else height_px(sheet, k) * 0.75 for k in range(r)) + y * 0.75
        want = {'セル': paste.col_name(c + 1) + str(r + 1), '左': left, '上': top, '幅': image['w'] * 0.75, '高さ': image['h'] * 0.75}
        have = {'セル': got['topLeft'], '左': got['left'], '上': got['top'], '幅': got['width'], '高さ': got['height']}
        bad = [k for k in want if (want[k] != have[k] if k == 'セル' else
                                   abs(want[k] - have[k]) > (POSITION_TOLERANCE if k in ('左', '上') else SIZE_TOLERANCE))]
        if bad:
            differ(f'画像 {i + 1}（{image["id"]}）の' + '・'.join(bad), want, have)
    return problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--xlsx', required=True, type=Path, help='build.py --xlsx で書き出したファイル')
    ap.add_argument('--input', type=Path, help='ブック JSON')
    ap.add_argument('--base', type=Path, help='このスキルで生成した HTML')
    ap.add_argument('--state', type=Path, help='ブラウザーで編集した後の状態テキスト（xlsx をこれで作ったとき）')
    ap.add_argument('--out', help='結果フォルダ（省略時は一時フォルダ）')
    args = ap.parse_args()
    if bool(args.input) == bool(args.base) or (args.state and not args.base):
        ap.error('--input か --base（xlsx を状態テキストで作ったなら --state も）を、build.py --xlsx と同じに指定してください。')
    if sys.platform != 'win32' or not shutil.which('pwsh') or not shutil.which('node'):
        print('この環境では検証できません（Windows、PowerShell 7 の pwsh、Node.js、デスクトップ版 Excel が必要）。', file=sys.stderr)
        return 2
    try:
        if args.input:
            workbook = build.read_json(args.input)
        else:
            workbook = build.apply_state(args.base, args.state) if args.state else build.read_base(args.base)
        build.validate(workbook)
    except (ValueError, KeyError, TypeError, OSError) as exc:
        print('エラー: ' + str(exc), file=sys.stderr)
        return 1
    if not args.xlsx.is_file():
        print(f'エラー: xlsx ファイルがありません: {args.xlsx}', file=sys.stderr)
        return 1

    out = Path(args.out) if args.out else Path(tempfile.mkdtemp(prefix='excel-xlsx-check-'))
    out.mkdir(parents=True, exist_ok=True)
    wb_path = out / 'workbook.json'
    wb_path.write_text(json.dumps(workbook, ensure_ascii=False), encoding='utf-8')
    pwsh = ['pwsh', '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', str(CHECK / 'excel.ps1')]
    summary, total = [], 0
    for sheet in workbook['sheets']:
        base = out / re.sub(r'[^\w-]', '_', sheet['id'])
        # The range and the notes on borders that two neighbouring cells give differently (a data problem Excel cannot
        # keep). The other notes are about pasting and do not apply to a file.
        info = json.loads(paste.run(['node', str(CHECK / 'clipboard.cjs'), str(wb_path), sheet['id'], paste.used_range(sheet), str(base) + '.range'], '範囲の計算'))
        meta = json.loads(Path(str(base) + '.range.notes.json').read_text(encoding='utf-8'))
        notes = {a: kind for a, kind in meta['notes'].items() if kind == 'border-conflict'}
        box = ['-R1', str(meta['r1']), '-C1', str(meta['c1']), '-R2', str(meta['r2']), '-C2', str(meta['c2'])]
        try:
            paste.run(pwsh + ['-Mode', 'xlsx', '-Workbook', str(wb_path), '-Sheet', sheet['id'], '-Xlsx', str(args.xlsx.resolve()), '-SheetName', sheet['name'],
                              '-Rows', str(sheet['rows']), '-Cols', str(sheet['cols']), '-Out', str(base)] + box, 'Excel での照合')
        except RuntimeError as exc:
            if 'Excel.Application' in str(exc) or '80040154' in str(exc):
                print('デスクトップ版 Excel を起動できません。この環境では検証できません。', file=sys.stderr)
                return 2
            raise
        unexpected, _, _ = paste.compare(str(base) + '.native.json', str(base) + '.xlsx.json', notes)
        cells = [d for d in unexpected if 'cause' not in d]
        conflicts = [d for d in unexpected if 'cause' in d]
        sheet_problems = check_sheet(workbook, sheet, json.loads(Path(str(base) + '.sheet.json').read_text(encoding='utf-8-sig')))
        total += len(cells) + len(sheet_problems)
        summary.append({'sheet': sheet['name'], 'range': info['range'], 'cells': cells, 'conflicts': conflicts, 'sheetProblems': sheet_problems})

    lines = ['# xlsx 検証', '', 'Excel（デスクトップ版）で、元の表を直接作った状態と、xlsx を開いた状態をセルごとに比べた。'
             '列幅・行の高さ・枠の固定・枠線・倍率・画像は、HTML の値（px × 0.75 = pt）と比べた。', '']
    for s in summary:
        lines += [f'## {s["sheet"]}（{s["range"]}）', '', f'- セルの想定外の差: {len(s["cells"])} 件', f'- シートの想定外の差: {len(s["sheetProblems"])} 件', '']
        if s['cells']:
            lines += ['| セル | 項目 | Excel で作った表 | xlsx |', '|---|---|---|---|']
            lines += [f'| {d["cell"]} | {d["label"]} | {d["excel"]!r} | {d["pasted"]!r} |' for d in s['cells'][:200]] + ['']
        if s['sheetProblems']:
            lines += ['| 項目 | HTML | xlsx を開いた Excel |', '|---|---|---|']
            lines += [f'| {d["item"]} | {d["page"]!r} | {d["excel"]!r} |' for d in s['sheetProblems'][:200]] + ['']
        if s['conflicts']:
            lines += ['隣り合うセルの同じ境界に、異なる罫線を指定している箇所がある（データの問題）。Excel はセルの辺ごとに線を1本しか持てない。', '']
            lines += [f'- {d["cell"]}（{d["label"]}）' for d in s['conflicts'][:50]] + ['']
    report = out / 'report.md'
    report.write_text('\n'.join(lines), encoding='utf-8')
    (out / 'report.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    for s in summary:
        print(f'{s["sheet"]} {s["range"]}: セルの想定外の差 {len(s["cells"])} 件 / シートの想定外の差 {len(s["sheetProblems"])} 件'
              + (f' / 罫線の食い違い {len(s["conflicts"])} 件' if s['conflicts'] else ''))
    print(f'報告: {report}')
    return 1 if total else 0


if __name__ == '__main__':
    sys.exit(main())
