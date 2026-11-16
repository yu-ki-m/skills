// Builds workbook HTML for the tests with the skill's own scripts/build.py, exactly as the skill does for users.
import {spawnSync} from 'node:child_process';
import {readFileSync, writeFileSync} from 'node:fs';
import {createRequire} from 'node:module';
import {join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {makeTempDir, removeTempDir} from './tempdirs.mjs';

/** The skill folder (the parent of tests/). */
export const SKILL_ROOT = fileURLToPath(new URL('../../', import.meta.url));
export const BUILD_SCRIPT = join(SKILL_ROOT, 'scripts', 'build.py');

/** assets/core.js, the same model code the page runs (for expected values). */
export const Core = createRequire(import.meta.url)(join(SKILL_ROOT, 'assets', 'core.js'));

let python;
/** Python 3.10+ as {command, args}: PYTHON, else python / python3 / py -3. null when none is found. */
export function findPython() {
  if (python !== undefined) return python;
  const candidates = [];
  if (process.env.PYTHON) candidates.push([process.env.PYTHON]);
  candidates.push(...(process.platform === 'win32' ? [['python'], ['py', '-3'], ['python3']] : [['python3'], ['python']]));
  python = null;
  for (const [command, ...args] of candidates) {
    const probe = spawnSync(command, [...args, '-c', 'import sys;print(sys.version_info[:2]>=(3,10))'], {encoding: 'utf8', timeout: 15000});
    if (probe.status === 0 && probe.stdout.trim() === 'True') { python = {command, args}; break; }
  }
  return python;
}

/** Runs scripts/build.py with args and returns {status, stdout, stderr}. */
export function runBuild(args, {cwd = SKILL_ROOT} = {}) {
  const py = findPython();
  if (!py) throw new Error('Python 3.10+ が見つかりません。環境変数 PYTHON に実行ファイルのパスを指定できます。');
  const result = spawnSync(py.command, [...py.args, BUILD_SCRIPT, ...args], {
    cwd, encoding: 'utf8', timeout: 60000, env: {...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1'},
  });
  if (result.error) throw result.error;
  return {status: result.status, stdout: result.stdout, stderr: result.stderr};
}

/** Writes book (an object) as JSON and builds it into dir/name. Returns the HTML path. */
export function buildBook(book, dir, name = 'book.html') {
  const json = join(dir, name.replace(/\.html$/i, '') + '.json');
  const html = join(dir, name);
  writeFileSync(json, JSON.stringify(book, null, 2), 'utf8');
  const result = runBuild(['--input', json, '--output', html]);
  if (result.status !== 0) throw new Error(`build.py が失敗しました（${name}）:\n${result.stderr || result.stdout}`);
  return html;
}

/** A new empty folder under the system temp folder. Remove it with removeDir() in an after() hook. */
export const tempDir = () => makeTempDir();
export const removeDir = removeTempDir;

/** A fresh copy of tests/fixtures/<name>.json, so that a test can change it freely. */
export function fixture(name) {
  return JSON.parse(readFileSync(join(SKILL_ROOT, 'tests', 'fixtures', name + '.json'), 'utf8'));
}

/** The workbook JSON embedded in a built HTML file. */
export function embeddedBook(htmlPath) {
  const match = /<script id="workbook-data" type="application\/json">(.*?)<\/script>/s.exec(readFileSync(htmlPath, 'utf8'));
  if (!match) throw new Error(`${htmlPath} に workbook-data がありません`);
  return JSON.parse(match[1]);
}

/** A minimal valid workbook with one sheet; patch overrides workbook fields, sheet overrides the sheet's. */
export function blankBook({sheet = {}, ...patch} = {}) {
  return {
    schema: 'excel-html-workbook/v1', id: 'test-blank', revision: 'draft', title: 'テスト', activeSheet: 'main',
    sheets: [{id: 'main', name: 'Sheet1', rows: 30, cols: 12, cells: {}, merges: [], rowHeights: [], colWidths: [], ...sheet}],
    ...patch,
  };
}
