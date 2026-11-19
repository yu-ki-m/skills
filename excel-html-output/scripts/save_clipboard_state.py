#!/usr/bin/env python3
"""Save the state text that 「AIに渡す」 put on the clipboard (「テキストをコピー」) to a file, as it was copied.

The text then need not be pasted into the chat nor typed out again: build.py reads the file (--state).
Only a whole state text is saved. Anything else on the clipboard is refused without showing any of it, since the
clipboard may hold something private. Line ends are saved as LF, as 「テキストをエクスポート」 writes them.
It also tells when the state was written out (and how long ago), and which HTML in the same folder it was made from:
the HTML itself, or the older version of it kept when the HTML was built again since (then the edits go to the HTML
by its name, with --output). Messages are written in UTF-8 whatever the console code page.

Exit codes: 0 saved (or the same as the file already saved), 1 refused (no state text on the clipboard, one cut
short, or the clipboard held by another program), 2 the clipboard of the screen cannot be read here (over SSH, in a
container, or no way to read it).
"""
import argparse
import base64
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

sys.dont_write_bytecode = True  # importing build.py must not leave a __pycache__ folder in the skill
sys.path.insert(0, str(Path(__file__).resolve().parent))
import build  # noqa: E402  (run_core reads the text with core.js; write_keeping_previous keeps an older file)

STATE_FIRST = "EXCEL_HTML_STATE_V1"
# The clipboard of Windows, from WSL: UTF-8 bytes in Base64, so that the console code page (CP932) changes no character.
PS_CLIPBOARD = ("$t = Get-Clipboard -Raw; if ($null -eq $t) { $t = '' }; "
                "[Console]::Write([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$t)))")


class Unavailable(Exception):
    """This environment has no clipboard of the screen that this script can read."""


def remote_reason():
    """Why the clipboard here is not the one of the screen the page is shown on, or None."""
    if os.environ.get("SSH_CONNECTION") or os.environ.get("SSH_CLIENT") or os.environ.get("SSH_TTY"):
        return "Claude Code が SSH でつないだ先で動いている"
    if os.environ.get("REMOTE_CONTAINERS") or os.environ.get("CODESPACES") or os.environ.get("DEVCONTAINER") or Path("/.dockerenv").exists():
        return "Claude Code がコンテナの中で動いている"
    return None


def in_wsl():
    if os.environ.get("WSL_DISTRO_NAME"):
        return True
    try:
        return "microsoft" in Path("/proc/version").read_text(encoding="utf-8", errors="replace").lower()
    except OSError:
        return False


# The tools that read the clipboard of Linux, with what each writes when it works but the clipboard holds no text
# (nothing copied, or only a picture): that is no text, not a clipboard that cannot be read here.
LINUX_TOOLS = (
    (["wl-paste", "--no-newline"], ("Nothing is copied", "No selection", "No suitable type of content copied")),
    (["xclip", "-selection", "clipboard", "-o"], ("target STRING not available", "target UTF8_STRING not available",
                                                  "There is no owner for the", "cannot convert")),
    (["xsel", "--clipboard", "--output"], ()),
)


def _win32():
    """user32 and kernel32, with the argument and result types the clipboard calls need: a HANDLE and a pointer are
    64 bits wide in a 64-bit Python, and ctypes would cut them to 32 bits without these."""
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL
    user32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
    user32.IsClipboardFormatAvailable.restype = wintypes.BOOL
    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype = wintypes.HANDLE
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    kernel32.GlobalSize.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalSize.restype = ctypes.c_size_t
    return ctypes, user32, kernel32


def _text_from_handle(handle):
    """The text of a CF_UNICODETEXT memory handle, up to its first NUL; None when it cannot be locked."""
    ctypes, _, kernel32 = _win32()
    pointer = kernel32.GlobalLock(handle)
    if not pointer:
        return None
    try:
        return ctypes.wstring_at(pointer, kernel32.GlobalSize(handle) // 2).split("\0", 1)[0]
    finally:
        kernel32.GlobalUnlock(handle)


def _windows_text():
    _, user32, _ = _win32()
    cf_unicodetext = 13
    # Another program may hold the clipboard open for a moment (a clipboard manager, the browser writing to it).
    for _ in range(20):
        if user32.OpenClipboard(None):
            break
        time.sleep(0.05)
    else:
        raise ValueError("クリップボードをほかのアプリが使っていて開けません。少し待ってから、もう一度実行してください。")
    try:
        if not user32.IsClipboardFormatAvailable(cf_unicodetext):
            return None
        handle = user32.GetClipboardData(cf_unicodetext)
        return _text_from_handle(handle) if handle else None
    finally:
        user32.CloseClipboard()


def _command_bytes(command, env=None, empty=()):
    """What command writes. When it fails with one of the messages in empty, the clipboard holds no text: b""."""
    try:
        result = subprocess.run(command, capture_output=True, check=False, timeout=30, env=env)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise Unavailable(f"{command[0]} でクリップボードを読めません（{exc}）") from exc
    if result.returncode:
        message = (result.stderr or b"").decode("utf-8", "replace")
        if any(part in message for part in empty):
            return b""
        raise Unavailable(f"{command[0]} でクリップボードを読めません（終了コード {result.returncode}）")
    return result.stdout


def _command_text(command, env=None):
    try:
        return _command_bytes(command, env).decode("utf-8")
    except UnicodeDecodeError as exc:
        raise Unavailable(f"{command[0]} が返した文字を UTF-8 として読めません") from exc


def _from_base64(data):
    try:
        return base64.b64decode(data.strip(), validate=True).decode("utf-8")
    except (ValueError, UnicodeDecodeError) as exc:
        raise Unavailable("PowerShell が返したクリップボードの文字を読めません") from exc


def _powershell_text():
    return _from_base64(_command_bytes(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", PS_CLIPBOARD]))


def _linux_text():
    """(text, where it was read) from the first of LINUX_TOOLS that works; the text is None when the clipboard holds
    no text (nothing, a picture, or bytes that are not UTF-8 text)."""
    tried = []
    for command, empty in LINUX_TOOLS:
        if not shutil.which(command[0]):
            continue
        try:
            data = _command_bytes(command, empty=empty)
        except Unavailable as exc:
            tried.append(str(exc))
            continue
        source = f"Linux のクリップボード（{command[0]}）"
        try:
            return (data.decode("utf-8") or None), source
        except UnicodeDecodeError:
            return None, source
    raise Unavailable("この環境ではクリップボードを読めません" + ("（" + " / ".join(tried) + "）" if tried else
                      "（wl-paste・xclip・xsel のどれもありません）"))


def read_clipboard_text():
    """(text, where it was read): the text on the clipboard of the screen, or None when it holds none, and the way it
    was read, as told to the person. Raises Unavailable when it is not the clipboard of the screen (SSH, a container)
    or cannot be read here."""
    reason = remote_reason()
    if reason:
        raise Unavailable(reason + "ので、画面を表示しているパソコンのクリップボードではないので、読みません")
    if sys.platform == "win32":
        return _windows_text(), "Windows のクリップボード"
    if sys.platform == "darwin":
        return _command_text(["pbpaste"], env={**os.environ, "LANG": "en_US.UTF-8", "LC_ALL": "en_US.UTF-8"}), "macOS のクリップボード（pbpaste）"
    if in_wsl() and shutil.which("powershell.exe"):  # the browser runs on Windows and writes to its clipboard
        return _powershell_text(), "WSL から Windows のクリップボード（powershell.exe）"
    return _linux_text()



def when(iso):
    """exportedAt as local time, with how long ago it was: 「2026-11-20 10:23:45（5分前。exportedAt …）」."""
    try:
        moment = datetime.fromisoformat(str(iso).replace("Z", "+00:00"))
        if moment.tzinfo is None:
            raise ValueError(iso)
    except ValueError:
        return f"不明（exportedAt {iso}）"
    minutes = round((datetime.now(timezone.utc) - moment).total_seconds() / 60)
    ago = ("この端末の時計より後" if minutes < 0 else f"{minutes}分前" if minutes < 60 else
           f"{minutes // 60}時間{minutes % 60}分前" if minutes < 24 * 60 else f"{minutes // (24 * 60)}日前")
    return f"{moment.astimezone().strftime('%Y-%m-%d %H:%M:%S')}（{ago}。exportedAt {iso}）"


def main(argv=None):
    build.utf8_output()
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("output", type=Path, help="保存先のファイル（例: 見積明細-edits.txt）。元のHTMLと同じフォルダーにする")
    args = parser.parse_args(argv)
    retry = "HTMLの「AIに渡す」で「テキストをコピー」を押してから、もう一度実行してください。"
    try:
        text, source = read_clipboard_text()
    except Unavailable as exc:
        print("エラー: " + str(exc) + "。「テキストをエクスポート」で保存したファイルを渡してもらってください。", file=sys.stderr)
        return 2
    except ValueError as exc:
        print("エラー: " + str(exc), file=sys.stderr)
        return 1
    if not text:
        print("エラー: クリップボードに文字がありません。" + retry, file=sys.stderr)
        return 1
    text = text.lstrip("﻿").replace("\r\n", "\n").replace("\r", "\n")
    if not text.startswith(STATE_FIRST):
        # Not a state text: say nothing of what it is.
        print("エラー: クリップボードの文字は、編集状態のテキストではありません（クリップボードの中身は表示しません）。" + retry, file=sys.stderr)
        return 1
    data = text.encode("utf-8")
    try:
        state = json.loads(build.run_core("process.stdout.write(JSON.stringify(C.readStateText(s)));", text))
        if not isinstance(state, dict) or state.get("schema") != "excel-html-state/v1" or not isinstance(state.get("workbook"), dict):
            raise ValueError("状態形式が不正です。")
        same = args.output.is_file() and args.output.read_bytes() == data
        previous = None if same else build.write_keeping_previous(args.output, data)
    except PermissionError as exc:
        print(f"エラー: 保存先のファイルを書き換えられません（{exc.filename}）。", file=sys.stderr)
        return 1
    except (ValueError, OSError) as exc:
        print("エラー: " + str(exc), file=sys.stderr)
        return 1
    workbook = state["workbook"]
    print(str(args.output.resolve()))
    if same:
        print("前回保存したものと同じ編集状態です（ファイルは書き換えていません）。"
              "画面で編集を続けたなら、その編集はまだコピーされていないので、もう一度「テキストをコピー」を押してもらってください。")
    if previous:
        print("以前の内容: " + str(previous.resolve()))
    print(f"題名: {workbook.get('title')} / ブックID: {state.get('workbookId')} / シート: {len(workbook.get('sheets') or [])}枚 / "
          f"書き出した時刻: {when(state.get('exportedAt'))}")
    print("読み取り元: " + source)
    # The HTML the state was made from, or, when that HTML was built again since the page was opened, the older
    # version kept of it: the edits then go to the HTML by its name (--output), or they would go into the older one.
    bases = build.matching_html(args.output.resolve().parent, state.get("workbookId"), state.get("baseRevision"))
    current = [path for path in bases if not build.older_of(path)]
    for path in current:
        print("元のHTML（ブックIDとリビジョンが一致）: " + str(path))
    for path in ([] if current else bases):
        original = build.older_of(path)
        print("以前の版（ブックIDとリビジョンが一致）: " + str(path))
        print(f"  今の {original.name} は、この状態の元になった版（ユーザーが画面を開いたときの版）の後に、作り直されています。"
              f"HTMLに反映するときは、--base にこの以前の版を、--output に {original} を指定してください。"
              f"JSONに取り出すときやxlsxにするときは、先にHTMLに反映してから、反映した後の {original.name} を --state なしで --base にしてください。")
    if not bases:
        print("同じフォルダーに、この状態の元のHTML（ブックIDとリビジョンが一致するもの）はありません。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
