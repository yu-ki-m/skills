// Helpers shared by the tests of scripts/build.py (tests/build*.test.mjs): the temp folder of the running test file,
// folders in it, the state text of 「AIに渡す」 for a built HTML, and the files a build leaves in a folder.
import {mkdirSync, readdirSync, readFileSync} from 'node:fs';
import {join} from 'node:path';
import {after, before} from 'node:test';
import {Core, embeddedBook, removeDir, tempDir} from './workbook.mjs';
import {readXlsx} from './xlsx.mjs';

/** The temp folder of the running test file (see useTempDir). */
export let dir;
/**
 * Registers before/after hooks for one test file: a temp folder is made before its first test and set as dir, and is
 * removed after its last. Each folder is kept apart, so all of them go even when several test files share a process.
 */
export function useTempDir() {
  let made;
  before(() => { dir = made = tempDir(); });
  after(() => removeDir(made));
}

/**
 * A state text of 「AIに渡す」 for the HTML at htmlPath, in format 'toon' (what the page gives by default) or 'json',
 * after edit(workbook, state) changes a copy of its workbook or the state around it.
 */
export function stateText(htmlPath, edit, format = 'toon') {
  const base = embeddedBook(htmlPath), workbook = structuredClone(base);
  const state = {schema: 'excel-html-state/v1', workbookId: base.id, baseRevision: base.revision, exportedAt: new Date().toISOString(), workbook};
  edit(workbook, state);
  return Core.stateText(state, format);
}
/** The folder name inside dir, made if it is not there yet (each test gives a name of its own, so tests share no files). */
export const newFolder = name => { const folder = join(dir, name); mkdirSync(folder, {recursive: true}); return folder; };
/** The names of the HTML files in folder, sorted. */
export const listHtml = folder => readdirSync(folder).filter(n => n.endsWith('.html')).sort();
/** The workbook read back from the xlsx file at path (tests/lib/xlsx.mjs). */
export const xlsxOf = path => readXlsx(readFileSync(path));
/** The kept older versions in folder (<stem>_<YYYYMMDDHHMMSS>.html) whose revision is revision. */
export const olderWith = (folder, revision) => listHtml(folder).filter(n => /_\d{14}\.html$/.test(n)).map(n => join(folder, n)).filter(p => embeddedBook(p).revision === revision);
