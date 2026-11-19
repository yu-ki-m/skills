// scripts/save_clipboard_state.py：「AIに渡す」の「テキストをコピー」でクリップボードに入った状態テキストを、そのままファイルに保存する。
// 利用者のクリップボードには触れない。スクリプトをモジュールとして読み込み、クリップボードを読む関数を差し替えて動かす。
import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {after, before, test} from 'node:test';
import {buildBook, Core, embeddedBook, findPython, fixture, removeDir, runBuild, SKILL_ROOT, tempDir} from './lib/workbook.mjs';

let dir;
before(() => { dir = tempDir(); });
after(() => removeDir(dir));
const newFolder = name => { const folder = join(dir, name); mkdirSync(folder, {recursive: true}); return folder; };

/**
 * Runs save_clipboard_state.py with output as its argument, the clipboard holding clipboard (a string, or null for a
 * clipboard without text), or, with unavailable, in an environment whose clipboard cannot be read.
 */
function save(output, clipboard, {unavailable = false, env = {PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1'}} = {}) {
  const fake = join(dir, 'clipboard-' + Math.random().toString(36).slice(2) + '.txt');
  if (clipboard !== null) writeFileSync(fake, clipboard, 'utf8');
  const code = ['import sys', 'sys.dont_write_bytecode = True', 'sys.path.insert(0, sys.argv[1])', 'import save_clipboard_state as m',
    'def read():',
    unavailable ? '    raise m.Unavailable("テストの環境ではクリップボードを読めません")'
      : '    return (open(sys.argv[2], encoding="utf-8", newline="").read() if sys.argv[2] else None), "テストのクリップボード"',
    'm.read_clipboard_text = read', 'sys.exit(m.main([sys.argv[3]]))'].join('\n');
  const py = findPython(), base = {...process.env};
  delete base.PYTHONIOENCODING;
  delete base.PYTHONUTF8;
  return spawnSync(py.command, [...py.args, '-c', code, join(SKILL_ROOT, 'scripts'), clipboard === null ? '' : fake, output],
    {encoding: 'utf8', timeout: 60000, env: {...base, ...env}});
}

/** The state text 「AIに渡す」 gives for the HTML at htmlPath after edit(workbook) changes a copy of its workbook. */
function stateText(htmlPath, edit, format = 'toon', exportedAt = '2026-11-20T01:23:45.678Z') {
  const base = embeddedBook(htmlPath), workbook = structuredClone(base);
  edit(workbook);
  return Core.stateText({schema: 'excel-html-state/v1', workbookId: base.id, baseRevision: base.revision, exportedAt, workbook}, format);
}

/** Runs Python code with the module imported as m (the clipboard itself is never read), and returns {status, stdout, stderr}. */
function python(code, env = {}) {
  const py = findPython();
  const program = ['import sys', 'sys.dont_write_bytecode = True', 'sys.path.insert(0, sys.argv[1])', 'import save_clipboard_state as m', code].join('\n');
  return spawnSync(py.command, [...py.args, '-c', program, join(SKILL_ROOT, 'scripts')],
    {encoding: 'utf8', timeout: 60000, env: {...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1', ...env}});
}

test('クリップボードの状態テキストを一字も変えずにファイルに保存し（改行はLFにする）、そのファイルで編集を反映できる', () => {
  for (const format of ['toon', 'json']) {
    const folder = newFolder('saved-' + format), html = buildBook(fixture('book'), folder, '表.html');
    const text = stateText(html, w => { w.sheets[0].cells.B4.v = 'クリップボードから'; }, format);
    const saved = join(folder, '表-edits.txt');
    // Windows のクリップボードの文字は、行の終わりが CRLF になる。
    const result = save(saved, text.replace(/\n/g, '\r\n'));
    assert.equal(result.status, 0, result.stderr);
    assert.equal(readFileSync(saved, 'utf8'), text, format + ': 中身は「テキストをエクスポート」と同じ');
    const lines = result.stdout.trim().split(/\r?\n/);
    assert.equal(lines[0], saved, '1行目は保存したファイル');
    assert.match(result.stdout, /題名: テスト用ブック \/ ブックID: /);
    assert.match(result.stdout, /exportedAt 2026-11-20T01:23:45\.678Z/, 'チャットに貼られたテキストと同じ書き出しかを、時刻で確かめられる');
    assert.ok(lines.includes('元のHTML（ブックIDとリビジョンが一致）: ' + html), '保存したフォルダーの中から、反映する先のHTMLを示す:\n' + result.stdout);
    const applied = runBuild(['--base', html, '--state', saved]);
    assert.equal(applied.status, 0, applied.stderr);
    assert.equal(embeddedBook(html).sheets[0].cells.B4.v, 'クリップボードから');
  }
});

test('状態テキストでないものは、中身を表示せずに止め、ファイルを作らない（空のときは、そう伝える）', () => {
  const folder = newFolder('not-state'), saved = join(folder, 'edits.txt');
  for (const [clipboard, reason] of [
    ['パスワード: hunter2-秘密', /クリップボードの文字は、編集状態のテキストではありません（クリップボードの中身は表示しません）/],
    [' EXCEL_HTML_STATE_V1 TOON\n先頭に空白', /クリップボードの文字は、編集状態のテキストではありません/],
    ['', /クリップボードに文字がありません/],
    [null, /クリップボードに文字がありません/],
  ]) {
    const result = save(saved, clipboard);
    assert.equal(result.status, 1, JSON.stringify(clipboard));
    assert.match(result.stderr, reason);
    assert.match(result.stderr, /「テキストをコピー」を押してから、もう一度実行してください/);
    assert.ok(!result.stderr.includes('hunter2') && !result.stdout.includes('hunter2'), '中身を一部も出さない');
    assert.equal(result.stdout, '');
    assert.deepEqual(readdirSync(folder), []);
  }
});

test('前回保存したものと同じ編集状態なら、ファイルを増やさずに、そう伝える（コピーし直し忘れに気づけるように）', () => {
  const folder = newFolder('same'), html = buildBook(fixture('book'), folder, '表.html'), saved = join(folder, 'edits.txt');
  const text = stateText(html, () => {});
  assert.equal(save(saved, text).status, 0);
  const again = save(saved, text);
  assert.equal(again.status, 0, again.stderr);
  assert.match(again.stdout, /前回保存したものと同じ編集状態です（ファイルは書き換えていません）/);
  assert.deepEqual(readdirSync(folder).filter(n => n.endsWith('.txt')), ['edits.txt'], '年月日時分秒を付けた複製を作らない');
});

test('書き出した時刻を、何分前かと、読み取った場所と一緒に示す', () => {
  const folder = newFolder('when'), html = buildBook(fixture('book'), folder, '表.html');
  const result = save(join(folder, 'edits.txt'), stateText(html, () => {}, 'toon', new Date(Date.now() - 5 * 60000).toISOString()));
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /（[45]分前。exportedAt \d{4}-\d\d-\d\dT/);
  assert.match(result.stdout, /^読み取り元: テストのクリップボード$/m, '実際に読んだ方法を示す');
});

test('SSH でつないだ先やコンテナの中では、画面のパソコンのクリップボードではないので、読まずに終了コード2で止まる', () => {
  // 読み取りの関数が呼ばれたら失敗にして、本物のクリップボードに触れないことも確かめる。
  const code = 'def never(*a, **k):\n    raise AssertionError("クリップボードを読んだ")\n' +
    'm._windows_text = m._command_text = m._command_bytes = m._linux_text = m._powershell_text = never\nsys.exit(m.main([sys.argv[2] if len(sys.argv) > 2 else "edits.txt"]))';
  for (const env of [{SSH_CONNECTION: '10.0.0.2 51000 10.0.0.1 22'}, {SSH_CLIENT: '10.0.0.2 51000 22'}, {SSH_TTY: '/dev/pts/0'}, {REMOTE_CONTAINERS: 'true'}, {CODESPACES: 'true'}, {DEVCONTAINER: 'true'}]) {
    const result = python(code, env);
    assert.equal(result.status, 2, JSON.stringify(env) + result.stderr);
    assert.match(result.stderr, /画面を表示しているパソコンのクリップボードではないので、読みません/);
    assert.match(result.stderr, /「テキストをエクスポート」で保存したファイルを渡してもらって/);
  }
});

test('WSL から読む Windows のクリップボードの文字は、文字コードに左右されない Base64 で受け取って戻す', () => {
  const result = python('import base64\n' +
    'text = "EXCEL_HTML_STATE_V1 TOON\\r\\n表😀é–€\\r\\n"\n' +
    'print(m._from_base64(base64.b64encode(text.encode("utf-8")) + b"\\r\\n") == text)\n' +
    'try:\n    m._from_base64(b"not base64!")\nexcept m.Unavailable as e:\n    print("Unavailable", e)');
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /^True\r?\n/);
  assert.match(result.stdout, /Unavailable PowerShell が返したクリップボードの文字を読めません/);
});

test('Windows のクリップボードのメモリから文字を読む処理は、4GBより上のアドレスのハンドルでも、最初の NUL までを読む', {skip: process.platform !== 'win32' && 'Windows だけ'}, () => {
  // クリップボードには触れず、同じ種類のメモリ（GlobalAlloc）を作って読ませる。
  const result = python([
    'import ctypes', 'from ctypes import wintypes',
    'k = ctypes.WinDLL("kernel32", use_last_error=True)',
    'k.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]; k.GlobalAlloc.restype = wintypes.HGLOBAL',
    'k.GlobalLock.argtypes = [wintypes.HGLOBAL]; k.GlobalLock.restype = ctypes.c_void_p',
    'k.GlobalUnlock.argtypes = [wintypes.HGLOBAL]; k.GlobalFree.argtypes = [wintypes.HGLOBAL]; k.GlobalFree.restype = wintypes.HGLOBAL',
    'text = "EXCEL_HTML_STATE_V1 TOON\\n表😀" * 2000',
    'data = (text + "\\0後ろのごみ").encode("utf-16-le") + b"\\0\\0"',
    'handles = [k.GlobalAlloc(0x0002, len(data)) for _ in range(40)]',
    'h = max(handles)',
    'p = k.GlobalLock(h); ctypes.memmove(p, data, len(data)); k.GlobalUnlock(h)',
    'print(m._text_from_handle(h) == text, h > 0xFFFFFFFF)',
    'for x in handles: k.GlobalFree(x)',
  ].join('\n'));
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /^True (True|False)/, result.stdout);
});

test('途中で切れた状態テキストは、理由を示して止め、ファイルを作らない', () => {
  const folder = newFolder('cut'), html = buildBook(fixture('book'), folder, '表.html'), saved = join(folder, 'edits.txt');
  const text = stateText(html, () => {}), lines = text.split('\n');
  const result = save(saved, lines.slice(0, Math.floor(lines.length / 2)).join('\n'));
  assert.equal(result.status, 1);
  assert.match(result.stderr, /EXCEL_HTML_STATE_END/);
  assert.ok(!existsSync(saved));
});

test('保存したフォルダーに元のHTMLがないときはそのことを、作り直した後のHTMLしかないときは以前の版と、反映するときの --output を示す', () => {
  const folder = newFolder('no-base'), html = buildBook(fixture('book'), folder, '表.html');
  const text = stateText(html, () => {});
  const other = newFolder('elsewhere'), saved = join(other, 'edits.txt');
  let result = save(saved, text);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /同じフォルダーに、この状態の元のHTML（ブックIDとリビジョンが一致するもの）はありません/);
  writeFileSync(join(folder, 'edits.txt'), stateText(html, w => { w.sheets[0].cells.B4.v = '改'; }), 'utf8');
  const rebuilt = runBuild(['--base', html, '--state', join(folder, 'edits.txt')]);
  assert.equal(rebuilt.status, 0, rebuilt.stderr);
  result = save(join(folder, 'edits2.txt'), text);
  assert.equal(result.status, 0, result.stderr);
  assert.ok(!/^元のHTML（/m.test(result.stdout), '作り直した表.html はリビジョンが違う:\n' + result.stdout);
  const older = result.stdout.trim().split(/\r?\n/).filter(l => l.startsWith('以前の版（ブックIDとリビジョンが一致）: '));
  assert.equal(older.length, 1, result.stdout);
  assert.match(older[0], /表_\d{14}\.html$/);
  assert.match(result.stdout, /今の 表\.html は、この状態の元になった版（ユーザーが画面を開いたときの版）の後に、作り直されています。/);
  assert.match(result.stdout, /HTMLに反映するときは、--base にこの以前の版を、--output に .*表\.html を指定してください。/);
  assert.match(result.stdout, /JSONに取り出すときやxlsxにするときは、先にHTMLに反映してから、反映した後の 表\.html を --state なしで --base にしてください。/);
});

test('名前に年月日時分秒が付いていても、元の名前のファイルが別の表なら、以前の版とは示さない', () => {
  const folder = newFolder('not-older'), dated = buildBook(fixture('book'), folder, '表_20260101000000.html');
  const other = fixture('book');
  other.id = 'other-book';
  buildBook(other, folder, '表.html');
  const result = save(join(folder, 'edits.txt'), stateText(dated, () => {}));
  assert.equal(result.status, 0, result.stderr);
  assert.ok(result.stdout.includes('元のHTML（ブックIDとリビジョンが一致）: ' + dated), result.stdout);
  assert.ok(!/以前の版/.test(result.stdout), result.stdout);
});

test('保存先に同じ名前のファイルがあると、年月日時分秒を付けた名前で残してから保存する', () => {
  const folder = newFolder('keep'), html = buildBook(fixture('book'), folder, '表.html'), saved = join(folder, 'edits.txt');
  writeFileSync(saved, '前に保存した状態', 'utf8');
  const result = save(saved, stateText(html, () => {}));
  assert.equal(result.status, 0, result.stderr);
  const kept = readdirSync(folder).filter(n => /^edits_\d{14}\.txt$/.test(n));
  assert.equal(kept.length, 1);
  assert.equal(readFileSync(join(folder, kept[0]), 'utf8'), '前に保存した状態');
  assert.match(result.stdout, /以前の内容: /);
});

test('クリップボードを読めない環境では、終了コード2で止まり、エクスポートしたファイルを頼むよう伝える', () => {
  const folder = newFolder('unavailable'), saved = join(folder, 'edits.txt');
  const result = save(saved, null, {unavailable: true});
  assert.equal(result.status, 2);
  assert.match(result.stderr, /テストの環境ではクリップボードを読めません/);
  assert.match(result.stderr, /「テキストをエクスポート」で保存したファイルを渡してもらって/);
  assert.ok(!existsSync(saved));
});

test('書き出した時刻は、何時間前・何日前・この端末の時計より後・読めない時刻も示す', () => {
  const result = python([
    'from datetime import datetime, timedelta, timezone',
    'now = datetime.now(timezone.utc)',
    'for d in (timedelta(minutes=90), timedelta(days=3), timedelta(minutes=-10)):',
    '    print(m.when((now - d).isoformat().replace("+00:00", "Z")))',
    'print(m.when("2026-11-20T01:23:45"))',
    'print(m.when("きのう"))',
  ].join('\n'));
  assert.equal(result.status, 0, result.stderr);
  const lines = result.stdout.trim().split(/\r?\n/);
  assert.match(lines[0], /（1時間30分前。exportedAt /);
  assert.match(lines[1], /（3日前。exportedAt /);
  assert.match(lines[2], /（この端末の時計より後。exportedAt /);
  assert.equal(lines[3], '不明（exportedAt 2026-11-20T01:23:45）', 'タイムゾーンのない時刻は、いつのことか決まらない');
  assert.equal(lines[4], '不明（exportedAt きのう）');
});

test('先頭に BOM のある状態テキストは BOM を除いて保存し、形式の違う状態テキストは保存しない', () => {
  const folder = newFolder('bom'), html = buildBook(fixture('book'), folder, '表.html'), saved = join(folder, 'edits.txt');
  const text = stateText(html, () => {});
  const result = save(saved, '\ufeff' + text);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(readFileSync(saved, 'utf8'), text);
  const base = embeddedBook(html), other = join(folder, 'other.txt');
  const wrong = Core.stateText({schema: 'excel-html-state/v2', workbookId: base.id, baseRevision: base.revision, exportedAt: '2026-11-20T01:23:45.678Z', workbook: base});
  const refused = save(other, wrong);
  assert.equal(refused.status, 1);
  assert.match(refused.stderr, /状態形式が不正です/);
  assert.ok(!existsSync(other));
});

test('表示は、端末の文字コードにかかわらず UTF-8 で書く（題名やパスに CP932 にない文字があっても、保存の後に落ちない）', () => {
  const folder = newFolder('utf8-é'), book = fixture('book');
  book.title = '見積 — é';
  const html = buildBook(book, folder, '表.html'), saved = join(folder, 'edits—é.txt');
  const result = save(saved, stateText(html, () => {}), {env: {}});
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /題名: 見積 — é \/ ブックID: /);
  assert.ok(result.stdout.includes(saved));
});

test('Linux で、クリップボードの道具は動いたが文字がないとき（空や画像だけ）は、読めない環境ではなく「文字がない」として扱う', () => {
  const result = python([
    'import subprocess',
    'def fake(stdout, stderr, code):',
    '    return lambda *a, **k: subprocess.CompletedProcess(a[0], code, stdout, stderr)',
    'm.shutil.which = lambda name: "/usr/bin/" + name if name == "wl-paste" else None',
    'm.subprocess.run = fake(b"", "Nothing is copied\\n".encode(), 1)',
    'print(m._linux_text())',
    'm.subprocess.run = fake(b"", "No suitable type of content copied\\n".encode(), 1)',
    'print(m._linux_text())',
    'm.subprocess.run = fake(b"\\x89PNG\\xff\\xfe", b"", 0)',
    'print(m._linux_text())',
    'm.subprocess.run = fake("EXCEL_HTML_STATE_V1 TOON\\n表".encode(), b"", 0)',
    'print(m._linux_text())',
    'm.shutil.which = lambda name: "/usr/bin/" + name if name == "xclip" else None',
    'm.subprocess.run = fake(b"", b"Error: target STRING not available\\n", 1)',
    'print(m._linux_text())',
    'm.subprocess.run = fake(b"", b"xclip: Error: There is no owner for the CLIPBOARD selection\\n", 1)',
    'print(m._linux_text())',
    'm.subprocess.run = fake(b"", b"xclip: Error: cannot convert CLIPBOARD selection to target STRING\\n", 1)',
    'print(m._linux_text())',
    'm.subprocess.run = fake(b"", b"Error: Can\\x27t open display: (null)\\n", 1)',
    'try:\n    m._linux_text()\nexcept m.Unavailable as e:\n    print("Unavailable", e)',
  ].join('\n'));
  assert.equal(result.status, 0, result.stderr);
  const lines = result.stdout.trim().split(/\r?\n/);
  assert.deepEqual(lines.slice(0, 5), [
    "(None, 'Linux のクリップボード（wl-paste）')", "(None, 'Linux のクリップボード（wl-paste）')", "(None, 'Linux のクリップボード（wl-paste）')",
    "('EXCEL_HTML_STATE_V1 TOON\\n表', 'Linux のクリップボード（wl-paste）')", "(None, 'Linux のクリップボード（xclip）')"]);
  assert.deepEqual(lines.slice(5, 7), ["(None, 'Linux のクリップボード（xclip）')", "(None, 'Linux のクリップボード（xclip）')"], 'xclip の新しい版の、文字がないときの文');
  assert.match(lines[7], /^Unavailable この環境ではクリップボードを読めません（xclip でクリップボードを読めません（終了コード 1））/);
});

test('WSL で powershell.exe が見つからないときは、Linux の道具で読み、読み取り元もその道具を示す', () => {
  const result = python([
    'import subprocess',
    'm.sys.platform = "linux"',
    'm.remote_reason = lambda: None',
    'm.in_wsl = lambda: True',
    'm.shutil.which = lambda name: "/usr/bin/wl-paste" if name == "wl-paste" else None',
    'm.subprocess.run = lambda *a, **k: subprocess.CompletedProcess(a[0], 0, "EXCEL_HTML_STATE_V1".encode(), b"")',
    'print(m.read_clipboard_text())',
  ].join('\n'));
  assert.equal(result.status, 0, result.stderr);
  assert.equal(result.stdout.trim(), "('EXCEL_HTML_STATE_V1', 'Linux のクリップボード（wl-paste）')");
});
