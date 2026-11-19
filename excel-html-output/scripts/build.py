#!/usr/bin/env python3
"""Build an offline HTML workbook or apply a version-matched state export, or write the workbook as an .xlsx file.

Python 3.10+ and Node.js 18+; no pip/npm packages. Validation uses the exact
same model validator as the browser, and a state export (TOON, the page's
default, or JSON) is read by the same code that wrote it (assets/core.js).
A book written by hand (--input) may use the shorthand (values, formulas,
styles), made into cells first. What Excel would not keep is found from the
book alone: border conflicts, texts too long for Excel and formulas the page
cannot read stop a book written by hand (unless --keep-user-content, for a
book taken out of the person's edits); formulas the page shows as #NAME?,
#REF! or #CIRC! are told with their cause. In anything else (a state, an
.xlsx file) they are what the person made, and are told.
The .xlsx file is written by scripts/xlsx.cjs from the same workbook.
Never executes input HTML or cell data.
An existing output file is never overwritten: it is first renamed to
<name>_<YYYYMMDDHHMMSS><suffix>, and the new content takes the original name.
Messages are written in UTF-8 whatever the console code page.
"""
import argparse
import hashlib
import html
import json
import os
import re
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA_RE = re.compile(r'<script id="workbook-data" type="application/json">(.*?)</script>', re.S)
# The name write_keeping_previous gives the older content of <name><suffix>.
OLDER_RE = re.compile(r"(.+)_\d{14}")
# What a sheet holds besides its cells, its merges and its sizes, compared when a build would undo a change.
SHEET_FIELDS = ("name", "rows", "cols", "freeze", "gridlines", "defaultStyle", "images")
KEEP_HINT = ("画面で編集した後の表やExcelのファイルから --extract-json で取り出したJSONで、ユーザーの内容のままにするなら、"
             "--keep-user-content を付けてください（止めずに「注意:」で知らせて作ります）。")


def utf8_output():
    """Writes stdout and stderr in UTF-8. Claude Code reads what a command writes as UTF-8, while Python on a Japanese
    Windows writes CP932 into a pipe (and fails on a character CP932 lacks, such as é in a path)."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
        except (AttributeError, ValueError):  # not a text stream that can be reconfigured
            pass


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def revision(workbook):
    data = {key: value for key, value in workbook.items() if key != "revision"}
    encoded = json.dumps(data, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return "sha256:" + hashlib.sha256(encoded).hexdigest()


def run_core(body, text):
    """Runs body (JavaScript, with C = assets/core.js and s = text) in Node and returns what it writes to stdout."""
    program = ("const C=require(process.argv[1]);let s='';process.stdin.setEncoding('utf8');process.stdin.on('data',x=>s+=x);"
               "process.stdin.on('end',()=>{try{" + body + "}catch(e){console.error(e.message);process.exitCode=1;}});")
    try:
        result = subprocess.run(["node", "-e", program, str(ROOT / "assets" / "core.js")],
                                input=text, text=True, encoding="utf-8", capture_output=True, check=False)
    except FileNotFoundError as exc:
        raise ValueError("Node.js 18+ が必要です。生成後のHTMLの閲覧にはNodeは不要です。") from exc
    if result.returncode:
        raise ValueError(result.stderr.strip())
    return result.stdout


def validate(workbook):
    run_core("C.validate(JSON.parse(s));", json.dumps(workbook, ensure_ascii=False))


def prepare(workbook, expand=False):
    """The workbook ready to build, checked as the browser checks it, and what Excel would not keep or show the same
    (one Node run). Returns (workbook, found).

    With expand (a book written by hand, --input), its shorthand (values, formulas, styles) is first made into cells
    by core.js (expand) and the expanded book is returned; any other book is returned as it was given.
    found holds conflicts (cell edges given two different lines, which Excel cannot keep: core.js borderConflicts, with
    the sheet name), notes (core.js expand: texts that read as numbers or dates) and what core.js cautions finds
    (limits, unreadable, uncomputable, dependent, errors, textFormulas). findings makes sentences of it.
    """
    body = ("const w=JSON.parse(s),notes=[],x=" + ("C.expand(w,n=>notes.push(n))" if expand else "w") + ";C.validate(x);"
            "const conflicts=[];for(const sh of x.sheets)for(const k of C.borderConflicts(x,sh))conflicts.push({sheet:sh.name,...k});"
            "process.stdout.write(JSON.stringify({workbook:" + ("x" if expand else "null") + ",conflicts,notes,...C.cautions(x)}));")
    found = json.loads(run_core(body, json.dumps(workbook, ensure_ascii=False)))
    return (found.pop("workbook") if expand else workbook), found


def findings(found, own, xlsx=False, kept=False):
    """The sentences of what prepare found: (problems, notes). A book Claude wrote (own) stops on what Excel cannot keep
    or the page cannot read, and is told how to mend the rest; in any other book (the person's edits, an .xlsx file)
    all of it is what the person made, and is told. A book taken out of those and mended (kept, --keep-user-content)
    may also hold what Claude added since, which is told to be mended if so. With xlsx, formulas the page cannot read
    are left to xlsx.cjs, which tells of each one it writes as a text."""
    problems, notes = [], list(found["notes"])
    stop = problems if own else notes
    if found["limits"]:
        stop.append(limit_message(found["limits"]) + ("取り出した後に自分で書いた文字なら、短くしてください。" if kept else ""))
    if found["unreadable"] and (own or not xlsx):
        stop.append("HTMLの数式として読めない式があります: " + listed(found["unreadable"]) +
                    "。名前・別のシートの参照・配列定数・閉じていないかっこや引用符などがある式で、"
                    "画面では #NAME? か #VALUE! になり、xlsx には文字列として入ります。" +
                    ("同じシートの番地と、HTMLで計算できる関数で書き直してください。" if own else
                     "取り出した後に自分で書いた式なら、書き直してください。" if kept else ""))
    if found["uncomputable"]:
        notes.append("HTMLでは計算できない数式があります（Excelでは計算されることがあるので、HTMLとExcelで表示が変わります）: " +
                     listed(found["uncomputable"]) + "。" +
                     ("関数は表にあるもので書き換えるか値を数値で入れ、シートの外の参照は rows・cols を増やすか参照を直し、"
                      "循環参照はなくしてください。" if own else ""))
    if found["dependent"]:
        notes.append("エラーのセルを参照していて、同じエラーになっている数式があります: " + listed(found["dependent"]) +
                     "。参照先のセルを直すと直ります。")
    if found["errors"]:
        notes.append("式の結果がエラーになる数式があります（Excelでも同じエラーになります）: " + listed(found["errors"]) + "。" +
                     ("INDEX などで範囲の外を指していないか、エラーの値を書いていないかを確かめてください。" if own else ""))
    if found["textFormulas"]:
        notes.append("表示形式が「@」（文字列）の数式のセルがあります: " + listed(found["textFormulas"]) +
                     "。画面やExcelでそのセルを編集すると、数式ではなく文字として入ります。" +
                     ("数式のセルには numFmt \"@\" を付けないでください。" if own else ""))
    if found["conflicts"]:
        stop.append(conflict_message(found["conflicts"], own, kept))
    return problems, notes


def listed(items, shown=10):
    """The first items, joined by 、, with how many more there are."""
    return "、".join(items[:shown]) + (f"、ほか{len(items) - shown}件" if len(items) > shown else "")


def limit_message(limits):
    return "Excelのセルに入らない長さがあります: " + listed(limits) + "。Excelに貼るかxlsxにすると、切れるか入りません。"


def read_input(path):
    """A book JSON written by hand (--input), its shorthand made into cells and checked. The verifiers read it so."""
    return prepare(read_json(path), expand=True)[0]


def embedded(path):
    """The workbook an HTML of this skill holds, or None: not a file, not UTF-8, no workbook data, or not a book."""
    try:
        match = DATA_RE.search(Path(path).read_text(encoding="utf-8-sig"))
        data = json.loads(match.group(1)) if match else None
    except (OSError, ValueError):  # UnicodeDecodeError is a ValueError
        return None
    return data if isinstance(data, dict) else None


def cell_order(address):
    """A cell address as (row, column), for sorting; an address that is not one goes last."""
    match = re.fullmatch(r"([A-Z]{1,3})(\d+)", address)
    if not match:
        return (sys.maxsize, 0)
    column = 0
    for letter in match.group(1):
        column = column * 26 + ord(letter) - 64
    return (int(match.group(2)), column)


def cell_name(row, col):
    """The A1 address of a 0-based row and column."""
    name = ""
    col += 1
    while col:
        col, rest = divmod(col - 1, 26)
        name = chr(65 + rest) + name
    return f"{name}{row + 1}"


def undone(base, current, result):
    """What writing result in place of current would undo, as (lost, both). current (the HTML to be replaced) was
    built again after base (the version the state was copied from), as when Claude built it from JSON while the person
    kept editing the page they had open. lost: what current changed from base that result has as base (cells, sheet
    fields, single column widths and row heights, merged cells, sheets added or removed, their order, the sheet opened
    first, the title); both: cells that the person changed too, which take the person's value."""
    if not current or current.get("id") != base.get("id") or current.get("revision") == base.get("revision"):
        return [], []
    def sheets(book):
        return {s.get("id"): s for s in book.get("sheets") or [] if isinstance(s, dict)}
    before, now_sheets, after, lost, both = sheets(base), sheets(current), sheets(result), [], []
    lost += [f"シート「{old.get('name')}」（作り直したときに消したもの）" for sheet_id, old in before.items()
             if sheet_id not in now_sheets and sheet_id in after]
    for sheet_id, now in now_sheets.items():
        old, new, name = before.get(sheet_id), after.get(sheet_id), now.get("name")
        if old is None:
            if new is None:
                lost.append(f"シート「{name}」（作り直したときに足したもの）")
            continue
        if new is None:
            continue
        cells, old_cells, new_cells = now.get("cells") or {}, old.get("cells") or {}, new.get("cells") or {}
        for a in sorted(set(cells) | set(old_cells), key=cell_order):
            if cells.get(a) != old_cells.get(a):
                if new_cells.get(a) == old_cells.get(a):
                    lost.append(f"{name}!{a}")
                elif new_cells.get(a) != cells.get(a):
                    both.append(f"{name}!{a}")
        lost += [f"シート「{name}」の {k}" for k in SHEET_FIELDS if now.get(k) != old.get(k) and new.get(k) == old.get(k)]
        for k in ("colWidths", "rowHeights"):
            sizes = [s.get(k) if isinstance(s.get(k), list) else [] for s in (now, old, new)]
            at = lambda values, i: values[i] if i < len(values) else None
            lost += [f"シート「{name}」の {k}[{i}]" for i in range(max(len(sizes[0]), len(sizes[1])))
                     if at(sizes[0], i) != at(sizes[1], i) and at(sizes[2], i) == at(sizes[1], i)]
        merges = [{json.dumps(m, sort_keys=True) for m in s.get("merges") or []} for s in (now, old, new)]
        label = lambda text: (lambda m: f"{cell_name(m['r1'], m['c1'])}:{cell_name(m['r2'], m['c2'])}")(json.loads(text))
        lost += [f"シート「{name}」の結合 {label(m)}（作り直したときに足したもの）" for m in sorted(merges[0] - merges[1]) if m not in merges[2]]
        lost += [f"シート「{name}」の結合 {label(m)}（作り直したときに外したもの）" for m in sorted(merges[1] - merges[0]) if m in merges[2]]
    order = [[s.get("id") for s in book.get("sheets") or [] if isinstance(s, dict)] for book in (current, base, result)]
    if order[0] != order[1] and order[2] == order[1]:
        lost.append("シートの順序")
    if current.get("activeSheet") != base.get("activeSheet") and result.get("activeSheet") == base.get("activeSheet"):
        lost.append("開くシート（activeSheet）")
    lost += [f"ブックの {k}" for k in ("title", "defaultStyle") if current.get(k) != base.get(k) and result.get(k) == base.get(k)]
    return lost, both


def older_of(path):
    """The file whose older content path is (write_keeping_previous names it <name>_<YYYYMMDDHHMMSS><suffix>), if
    that file is there and holds the same book (a file only named so is not an older version)."""
    path = Path(path)
    match = OLDER_RE.fullmatch(path.stem)
    original = path.with_name(match.group(1) + path.suffix) if match else None
    if not original or not original.is_file():
        return None
    book, older = embedded(original), embedded(path)
    return original if book and older and book.get("id") == older.get("id") else None


def matching_html(folder, workbook_id, revision):
    """The HTML files in folder holding the book a state was made from (the same book ID and revision)."""
    found = []
    for path in sorted(Path(folder).glob("*.html")):
        data = embedded(path)
        if data and data.get("id") == workbook_id and data.get("revision") == revision:
            found.append(path)
    return found


def already_applied(path, mode):
    """The sentence for a state that the HTML at path already holds."""
    if mode == "html":
        return (f"この編集状態は、すでに {Path(path).resolve()} に反映されています（反映しても内容が変わりません）。"
                "画面で編集を続けたなら、その編集はまだ渡されていません。"
                "画面を再読み込みせずに、もう一度「テキストをコピー」を押してから渡してもらってください。")
    return (f"この編集状態は、すでに --base の {Path(path).resolve()} に反映されています。"
            "--state を付けずに、--base だけで作ってください（同じ表になります）。")


def conflict_message(conflicts, own=True, kept=False, shown=10):
    """A sentence on the border conflicts: how many, the first few edges with their two lines, and, in a book Claude
    wrote (own), how to mend them; lines the person drew are kept as they are, and in a book taken out of the person's
    edits and mended (kept), lines added since are to be mended."""
    edges = []
    for k in conflicts:
        sides = ("下", "上") if k["edge"] == "下/上" else ("右", "左")
        edge = f'{k["sheet"]}!{k["cells"][0]} の{sides[0]}と {k["cells"][1]} の{sides[1]}（{k["lines"][0]} / {k["lines"][1]}）'
        if edge not in edges:  # a merged cell meets a neighbour along several cells
            edges.append(edge)
    more = f"、ほか{len(edges) - shown}か所" if len(edges) > shown else ""
    return (f"罫線の食い違いが{len(edges)}か所あります（Excelはセルの辺ごとに線を1本しか持てないので、"
            f"Excelに貼るかxlsxにすると、どちらか一方の線になります）: {'、'.join(edges[:shown])}{more}。"
            "画面では、上または左のセルの線（各組の1本目）を描いています。" +
            ("片方のセルにだけ書くか、同じ線にそろえてください（省略形の styles で引いた線は、後に書いた線にそろいます）。"
             "見た目を変えずにそろえるなら、下または右のセルの向かい合う辺の線を消してください。" if own else
             "ユーザーが付けた線（画面での編集や、元のxlsxの線）なら、そのまま残しました。取り出した後に自分で加えた線なら、直してください。" if kept else
             "ユーザーが付けた線（画面での編集や、元のxlsxの線）なので、そのまま残しました。"))


def read_state(path):
    """The state object of a 「AIに渡す」 text, TOON or JSON. A text cut short or changed inside is refused."""
    text = Path(path).read_text(encoding="utf-8-sig")
    return json.loads(run_core("process.stdout.write(JSON.stringify(C.readStateText(s)));", text))


def to_xlsx(workbook):
    """The .xlsx file of workbook (bytes, from scripts/xlsx.cjs) and its cautions, each a line starting with 注意:."""
    try:
        result = subprocess.run(["node", str(ROOT / "scripts" / "xlsx.cjs")], input=json.dumps(workbook, ensure_ascii=False).encode("utf-8"),
                                capture_output=True, check=False)
    except FileNotFoundError as exc:
        raise ValueError("Node.js 18+ が必要です。") from exc
    messages = result.stderr.decode("utf-8", "replace").strip()
    if result.returncode:
        raise ValueError(messages)
    return result.stdout, [line for line in messages.splitlines() if line.strip()]


def from_xlsx(path):
    """The workbook read from an .xlsx file (scripts/read_xlsx.cjs) and the sentences on what could not be carried over."""
    try:
        result = subprocess.run(["node", str(ROOT / "scripts" / "read_xlsx.cjs"), str(path)], capture_output=True, check=False)
    except FileNotFoundError as exc:
        raise ValueError("Node.js 18+ が必要です。") from exc
    if result.returncode:
        raise ValueError(result.stderr.decode("utf-8", "replace").strip())
    data = json.loads(result.stdout.decode("utf-8"))
    return data["workbook"], data["warnings"]


def read_base(path):
    source = Path(path).read_text(encoding="utf-8-sig")
    match = DATA_RE.search(source)
    if not match:
        raise ValueError("元HTMLにworkbook-dataがありません。このスキルの生成HTMLを指定してください。")
    base = json.loads(match.group(1))
    validate(base)
    if base["revision"] != revision(base):
        raise ValueError("元HTMLのデータがリビジョン生成後に変更されています。元の生成版を指定してください。")
    return base


# What to do when a state was copied from an older version of the HTML, by what is written: the edits go into the
# HTML the person opens (html), or only the older version changes in the command (an .xlsx file, the JSON, a check).
OLDER_BASE = {
    "html": "リビジョンが一致するものを --base に、元のファイル名を --output に指定してください"
            "（元のHTMLにだけある、作り直したときの変更は元に戻り、そのセルは「注意:」に出ます）。",
    "xlsx": "リビジョンが一致するものを --base に指定してください（--xlsx はそのまま）。"
            "ただし、作り直した後の変更はxlsxに入らない（「注意:」に出ます）ので、先にこの編集状態をHTMLに反映し"
            "（--base に以前の版、--output に元のファイル名）、反映した後のHTMLから作るほうが確かです。"
            "この編集状態をすでにHTMLに反映したなら、--state を付けずに、反映した後のHTMLを --base にしてください。",
    "json": "リビジョンが一致するものを --base に指定してください（--output のJSONはそのまま）。"
            "ただし、作り直した後の変更はJSONに入らない（「注意:」に出ます）ので、先にこの編集状態をHTMLに反映し"
            "（--base に以前の版、--output に元のファイル名）、反映した後のHTMLから取り出すほうが確かです。"
            "この編集状態をすでにHTMLに反映したなら、--state を付けずに、反映した後のHTMLを --base にして取り出してください。",
    "verify": "リビジョンが一致するものを --base に指定してください（ほかの引数は、xlsxを書き出したときと同じにする）。",
}


def apply_state(base_path, state_path, mode="html"):
    """The workbook of the state, checked against the HTML it was copied from (base_path). mode (OLDER_BASE) says what
    the state goes into, which changes what to do when it was copied from an older version."""
    base = read_base(base_path)
    state = read_state(state_path)
    if not isinstance(state, dict) or state.get("schema") != "excel-html-state/v1":
        raise ValueError("状態形式が不正です。")
    if state.get("workbookId") != base["id"]:
        title = (state.get("workbook") or {}).get("title") if isinstance(state.get("workbook"), dict) else None
        raise ValueError(f"この編集状態は別の表のものです（状態の題名: {title}、ブックID: {state.get('workbookId')}。"
                         f"元のHTMLのブックID: {base['id']}）。どの表の編集状態かを確かめてください。")
    if state.get("baseRevision") != base["revision"]:
        path = Path(base_path)
        if isinstance(state.get("workbook"), dict) and revision(state["workbook"]) == base["revision"]:
            raise ValueError(already_applied(path, mode))  # the HTML was built from this very state
        original = older_of(path) or path
        found = [p.name for p in matching_html(path.resolve().parent, base["id"], state.get("baseRevision"))]
        raise ValueError("この編集状態は、--base のHTMLの版ではなく、別の版（以前の版など）から書き出されたものです"
                         "（作り直した後のHTMLに、作り直す前の画面で編集した状態を渡したときなど）。"
                         f"同じフォルダーの以前の版（{original.stem}_<年月日時分秒>{original.suffix}）から、" + OLDER_BASE[mode] +
                         (f"リビジョンが一致するのは {'、'.join(found)} です。" if found else "同じフォルダーには、リビジョンが一致するHTMLはありません。"))
    workbook = state.get("workbook")
    validate(workbook)
    if workbook["id"] != base["id"] or workbook["revision"] != base["revision"]:
        raise ValueError("編集状態内のブックIDまたはリビジョンが一致しません。")
    return workbook


def build(workbook):
    """The HTML of workbook, which prepare has checked."""
    workbook["revision"] = revision(workbook)
    data = json.dumps(workbook, ensure_ascii=False, separators=(",", ":"))
    # JSON stays parseable but cannot terminate its inert script element.
    data = data.replace("<", "\\u003c").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    tokens = {
        "__TITLE__": html.escape(workbook["title"]),
        "/*__STYLE__*/": (ROOT / "assets" / "style.css").read_text(encoding="utf-8"),
        "/*__TOON__*/": (ROOT / "assets" / "toon.js").read_text(encoding="utf-8"),
        "/*__FORMULAS__*/": (ROOT / "assets" / "formulas.js").read_text(encoding="utf-8"),
        "/*__CORE__*/": (ROOT / "assets" / "core.js").read_text(encoding="utf-8"),
        "/*__APP__*/": (ROOT / "assets" / "app.js").read_text(encoding="utf-8"),
        "/*__DATA__*/": data,
    }
    source = (ROOT / "assets" / "shell.html").read_text(encoding="utf-8")
    # One pass prevents user text which resembles a token from being replaced.
    return re.sub("|".join(re.escape(token) for token in tokens), lambda m: tokens[m.group()], source)


def previous_name(path):
    """The name that keeps the old content: <stem>_<YYYYMMDDHHMMSS><suffix>, unused in the same folder."""
    moment = datetime.now()
    while True:
        candidate = path.with_name(f"{path.stem}_{moment:%Y%m%d%H%M%S}{path.suffix}")
        if not candidate.exists():
            return candidate
        moment += timedelta(seconds=1)


def write_keeping_previous(path, content):
    """Writes content (text, or bytes) to path. If path exists, it is renamed to previous_name(path) first.

    The new content is written to a temporary file beforehand, so a failure leaves the old file in place.
    Returns the path of the renamed old file, or None when there was nothing to keep.
    """
    if path.is_dir():
        raise ValueError(f"出力先がフォルダーです: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    if isinstance(content, bytes):
        temp.write_bytes(content)
    else:
        temp.write_text(content, encoding="utf-8")
    previous = None
    try:
        if path.exists():
            previous = previous_name(path)
            path.rename(previous)
        temp.replace(path)
    except OSError:
        if previous and previous.exists() and not path.exists():
            previous.rename(path)
        temp.unlink(missing_ok=True)
        raise
    return previous


def main():
    utf8_output()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path,
                        help="ブックJSON（手で書いたもの。範囲でまとめて書く省略形を使える。--extract-json で取り出したものも）")
    parser.add_argument("--base", type=Path, help="編集元の生成HTML")
    parser.add_argument("--state", type=Path, help="AI用状態テキスト（TOONかJSON）")
    parser.add_argument("--output", type=Path,
                        help="出力先。--base と --state で反映するときは省略でき、省略すると --base と同じファイルに書き出す"
                             "（--base が以前の版のときは省略できない）")
    parser.add_argument("--extract-json", action="store_true",
                        help="HTMLの代わりに、ビルドするブックをJSONで --output に書き出す。--input なら省略形を広げた後、"
                             "--base なら --state を反映した後（--state がなければそのHTMLの表）、--from-xlsx なら読み込んだもの。"
                             "表を直すときや、広げた形を見るときに使う")
    parser.add_argument("--xlsx", type=Path,
                        help="Excelのファイル（.xlsx）をここに書き出す（HTMLは書き出さない）。--input、--base（生成したHTMLの表）、"
                             "--base と --state（ブラウザーで編集した後の表）のどれかと使う")
    parser.add_argument("--from-xlsx", type=Path,
                        help="Excelのファイル（.xlsx）を読み込んで、--output にHTMLを作る（--extract-json ならJSON）")
    parser.add_argument("--keep-user-content", action="store_true",
                        help="--input のブックに、罫線の食い違い・Excelのセルに入らない長さ・HTMLで読めない数式があっても止めずに、"
                             "「注意:」で知らせて作る（画面で編集した後の表やExcelのファイルから --extract-json で取り出したJSONを直して作るとき。"
                             "自分で書いたブックには付けない）")
    args = parser.parse_args()
    if args.from_xlsx:
        if args.input or args.base or args.state or args.xlsx:
            parser.error("--from-xlsx は --input・--base・--state・--xlsx と一緒には使えません。")
        if not args.output:
            parser.error("--from-xlsx には --output を指定してください。")
        output_path = args.output
    elif args.xlsx:
        if args.output or args.extract_json:
            parser.error("--xlsx は --output・--extract-json と一緒には使えません。")
        if bool(args.input) == bool(args.base) or (args.state and not args.base):
            parser.error("--xlsx には、--input か --base（ブラウザーで編集した後の表なら --state も）を指定してください。")
        output_path = args.xlsx
    else:
        if bool(args.input) == bool(args.base) or (args.state and not args.base) or (args.base and not args.state and not args.extract_json):
            parser.error("--input、または --base と --state の組を指定してください（--extract-json なら --base だけでもよい）。")
        output_path = args.output or (args.base if args.base and args.state and not args.extract_json else None)
        if output_path is None:
            parser.error("--output を指定してください。")
    mode = "xlsx" if args.xlsx else "json" if args.extract_json else "html"
    # A book Claude wrote is mended before it is handed over; in anything else, what is found is what the person made.
    own = bool(args.input) and not args.keep_user_content
    try:
        if args.state and mode == "html" and not args.output and older_of(args.base):
            original = older_of(args.base)
            raise ValueError(f"--base の {args.base.name} は、{original.name} の以前の版です。--output を省くと、この以前の版に書き出すので、"
                             f"ユーザーが開く {original.name} は変わりません。--output に {original.resolve()} を指定してください"
                             "（以前の版のファイルに書き出すときも、--output にそのファイルを指定してください）。")
        if mode == "json" and embedded(output_path) is not None:
            raise ValueError(f"--output の {output_path.name} は、このスキルのHTMLです。--extract-json はJSONを書き出すので、"
                             f"HTMLがJSONに置き換わります。--output にはJSONのファイル（例 {output_path.with_name(output_path.stem + '-edit.json')}）"
                             "を指定してください。")
        if args.from_xlsx:
            workbook, cautions = from_xlsx(args.from_xlsx)
            for caution in cautions:
                print("注意: " + caution, file=sys.stderr)
        elif args.input:
            workbook = read_json(args.input)
        else:
            workbook = apply_state(args.base, args.state, mode) if args.state else read_base(args.base)
        # The HTML the state's book replaces (html), or, when a state copied from an older version goes into an .xlsx
        # file or JSON, the HTML built since by the older version's name: what it changed is not in the state.
        replaced = output_path if mode == "html" else older_of(args.base) if args.state else None
        current = embedded(replaced) if args.state and replaced else None
        if mode == "html" and current is not None:
            if current.get("id") != workbook["id"]:
                raise ValueError(f"--output の {output_path.name} は、別の表のHTMLです（ブックID: {current.get('id')}、題名: {current.get('title')}）。"
                                 "反映すると、その表が置き換わります。反映する先のHTMLを確かめるか、別の名前に書き出してください。")
            # A state given again (copied before more edits were made, and not copied since): applying it would
            # change nothing, while the page still holds the later edits.
            if current.get("revision") == revision(workbook):
                raise ValueError(already_applied(output_path, mode))
        workbook, found = prepare(workbook, expand=bool(args.input))
        problems, notes = findings(found, own, xlsx=bool(args.xlsx), kept=bool(args.input) and args.keep_user_content)
        if current is not None:
            lost, both = undone(embedded(args.base), current, workbook)
            if lost and mode == "html":
                notes.append(f"この編集状態の元になった版（ユーザーが画面を開いたときの版）の後に、{replaced.name} が作り直されています。"
                             f"その作り直しで変わったところは、反映すると元に戻ります: {listed(lost)}。"
                             "手順3などで自分（Claude）が加えた変更なら、反映した後にもう一度加えてください。"
                             "前に反映した編集状態の変更なら、ユーザーが画面で戻したものなので、そのままにしてください"
                             "（書き換える前の内容は「以前の内容:」のファイルに残ります）。")
            elif lost:
                notes.append(f"この編集状態の元になった版の後に、{replaced.name} が作り直されています。"
                             f"その作り直しで変わったところは、この{'xlsx' if args.xlsx else 'JSON'}には入っていません: {listed(lost)}。"
                             f"先にこの編集状態を {replaced.name} に反映すると（--base に以前の版、--output に {replaced.name}）、"
                             "元に戻るところを確かめてから、反映した後のHTMLから作れます。")
            if both:
                notes.append(f"画面の編集と作り直しの両方で変わったセルがあり、ユーザーの値にしました: {listed(both)}。"
                             "作り直したときの値が要るなら、ユーザーに確かめてください。")
        for note in notes:
            print("注意: " + note, file=sys.stderr)
        if problems:
            raise ValueError("ブック形式エラー: " + "".join(problems) + KEEP_HINT)
        if args.xlsx:
            output, cautions = to_xlsx(workbook)
            for caution in cautions:
                print(caution, file=sys.stderr)
        else:
            output = json.dumps(workbook, ensure_ascii=False, indent=2) if args.extract_json else build(workbook)
        previous = write_keeping_previous(output_path, output)
        print(str(output_path.resolve()))
        if previous:
            print("以前の内容: " + str(previous.resolve()))
    except PermissionError as exc:
        print(f"エラー: 書き出し先のファイルを書き換えられません。Excelなどで開いているなら、閉じてからやり直してください（{exc.filename}）。",
              file=sys.stderr)
        return 1
    except (ValueError, KeyError, TypeError, OSError) as exc:
        print("エラー: " + str(exc), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
