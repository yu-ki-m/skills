#!/usr/bin/env node
// excel-html-output の自動テストを、まとめて実行する。
//
//   node tests/run.mjs                          すべてのテスト
//   node tests/run.mjs ui-edit build            ファイル名に ui-edit か build を含むテストファイルだけ
//   node tests/run.mjs --test-name-pattern=結合  名前に「結合」を含むテストだけ（-- で始まる引数は node --test に渡す）
//   EXCEL_HTML_TEST_HEADED=1 node tests/run.mjs ブラウザーの画面を表示して実行する（操作を目で追うとき）
//
// 必要なもの: Node.js 22+、Python 3.10+（PYTHON で指定可）、Chrome か Edge（EXCEL_HTML_TEST_BROWSER で指定可）。
// npm や pip のパッケージは使わない。利用者のクリップボードには触れない。
// 終了コード: 0 = すべて合格（既知の不具合として記録した todo のテストは数えない）、1 = 失敗あり、2 = この環境では実行できない。
import {spawnSync} from 'node:child_process';
import {readdirSync} from 'node:fs';
import {availableParallelism} from 'node:os';
import {join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {findBrowser} from './lib/browser.mjs';
import {removeStale} from './lib/tempdirs.mjs';
import {findPython} from './lib/workbook.mjs';

const here = fileURLToPath(new URL('.', import.meta.url));
const problems = [];
if (Number(process.versions.node.split('.')[0]) < 22 || typeof WebSocket !== 'function') problems.push(`Node.js 22 以上が必要です（現在 ${process.version}）`);
if (!findPython()) problems.push('Python 3.10 以上が見つかりません（環境変数 PYTHON に実行ファイルを指定できます）');
if (!findBrowser()) problems.push('Chrome または Edge が見つかりません（環境変数 EXCEL_HTML_TEST_BROWSER に実行ファイルを指定できます）');
if (problems.length) {
  console.error('excel-html-output のテストを、この環境では実行できません。\n' + problems.map(p => '- ' + p).join('\n'));
  process.exit(2);
}

removeStale();
const args = process.argv.slice(2);
const filters = args.filter(a => !a.startsWith('--')), options = args.filter(a => a.startsWith('--'));
const files = readdirSync(here).filter(name => /\.test\.[cm]?js$/.test(name))
  .filter(name => !filters.length || filters.some(f => name.includes(f))).sort().map(name => join(here, name));
if (!files.length) {
  console.error('該当するテストファイルがありません: ' + filters.join(' '));
  process.exit(2);
}
// Each file runs in its own process with its own browser, which draws on the CPU. Running more files at once than
// there are spare CPUs makes every browser slower (measured on 8 logical CPUs: 4 files 49s, 6 files 43s, 8 files 54s).
const concurrency = Math.max(1, Math.min(6, availableParallelism() - 2));
const result = spawnSync(process.execPath, ['--test', '--test-reporter=spec', `--test-concurrency=${concurrency}`, '--test-timeout=120000', ...options, ...files], {stdio: 'inherit'});
process.exit(result.status ?? 1);
