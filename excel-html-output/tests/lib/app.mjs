// Helpers that operate a built workbook HTML the way a person does (click, type, keys, ribbon, dialogs) and read
// back what the person would see (cells, name box, formula bar, status bar) or hand to the AI (the state text).
//
// The tests never touch the person's real clipboard: an in-page clipboard replaces navigator.clipboard, the
// legacy document.execCommand('copy' | 'cut' | 'paste') is refused, and copy / cut / paste are sent as the same
// ClipboardEvent the browser fires for Ctrl+C / Ctrl+X / Ctrl+V, with a DataTransfer the test can read.
import {after, before} from 'node:test';
import {pathToFileURL} from 'node:url';
import {launch} from './browser.mjs';
import {buildBook, removeDir, tempDir} from './workbook.mjs';

const INIT_SCRIPT = `(() => {
  const store = {items: [], writes: 0, legacy: [], denyRead: false, denyWrite: false};
  const fake = {
    async write(items) {
      if (store.denyWrite) throw new DOMException('書き込みが許可されていません', 'NotAllowedError');
      const entries = [];
      for (const item of items) for (const type of item.types) entries.push({type, blob: await item.getType(type)});
      store.items = entries; store.writes++;
    },
    async writeText(text) {
      if (store.denyWrite) throw new DOMException('書き込みが許可されていません', 'NotAllowedError');
      store.items = [{type: 'text/plain', blob: new Blob([String(text)], {type: 'text/plain'})}]; store.writes++;
    },
    async read() {
      if (store.denyRead) throw new DOMException('読み取りが許可されていません', 'NotAllowedError');
      return store.items.length ? [new ClipboardItem(Object.fromEntries(store.items.map(e => [e.type, e.blob])))] : [];
    },
    async readText() { const e = store.items.find(x => x.type === 'text/plain'); return e ? e.blob.text() : ''; },
  };
  Object.defineProperty(Navigator.prototype, 'clipboard', {configurable: true, get: () => fake});
  const exec = Document.prototype.execCommand;
  Document.prototype.execCommand = function(command, ...rest) {
    if (/^(copy|cut|paste)$/i.test(command)) { store.legacy.push(command); return false; }
    return exec.call(this, command, ...rest);
  };
  window.__opened = [];
  window.open = (...args) => { window.__opened.push(args.map(String)); return null; };
  // Ctrl/Cmd + C, V, X: record whether the page left the key to the browser (it must, for copy and paste events to
  // fire), then stop the browser's own clipboard command. This listener on window runs after the page's own ones.
  window.__clipboardKeys = [];
  window.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && /^[cvx]$/i.test(e.key)) { window.__clipboardKeys.push({key: e.key.toLowerCase(), leftToBrowser: !e.defaultPrevented}); e.preventDefault(); }
  });
  window.__testClipboard = {
    async dump() {
      return Promise.all(store.items.map(async e => ({type: e.type, size: e.blob.size, text: e.type.startsWith('text/') ? await e.blob.text() : null})));
    },
    set(entries) { store.items = entries.map(e => ({type: e.type, blob: new Blob([e.text], {type: e.type})})); },
    get writes() { return store.writes; },
    get legacy() { return store.legacy.slice(); },
    denyRead(value) { store.denyRead = value; },
    denyWrite(value) { store.denyWrite = value; },
  };
})();`;

/** '#e2efda' -> 'rgb(226, 239, 218)', the form getComputedStyle returns. */
export const rgb = hex => 'rgb(' + [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16)).join(', ') + ')';

/**
 * Registers before/after hooks for one test file: one browser, and the workbooks built on demand into a temp folder.
 * open(t, book) opens the workbook for test t (book is a workbook object or a built HTML path) in a new document.
 * When t ends, the page must not have reported an error; its tab is then reused by a later test, which loads a new
 * document into it (opening a tab costs more than loading a page, and a new document shares nothing with the old).
 */
export function useBrowser() {
  const env = {browser: null, dir: null, built: new Map(), idle: []};
  before(async () => {
    env.dir = tempDir();
    env.browser = await launch();
  }, {timeout: 60000});
  after(async () => {
    await env.browser?.close();
    removeDir(env.dir);
  }, {timeout: 30000});
  env.build = book => {
    const key = JSON.stringify(book);
    if (!env.built.has(key)) env.built.set(key, buildBook(book, env.dir, `book-${env.built.size + 1}.html`));
    return env.built.get(key);
  };
  env.open = async (t, book) => {
    const html = typeof book === 'string' ? book : env.build(book);
    let page = env.idle.pop();
    if (page) {
      // Leave without the page's "unsaved edits" prompt: this capture listener runs before the page's own one.
      await page.evaluate(() => { window.addEventListener('beforeunload', e => e.stopImmediatePropagation(), true); }).catch(() => {});
    } else {
      page = await env.browser.newPage();
      await page.addInitScript(INIT_SCRIPT);
    }
    page.errors.length = 0;
    await page.send('Page.bringToFront');
    await page.goto(pathToFileURL(html).href);
    const sheet = new ExcelPage(page, html);
    t.after(() => sheet.release(env.idle));
    await page.waitFor(() => document.getElementById('grid')?.tBodies[0]?.rows.length > 0 && document.activeElement?.id === 'cell-editor',
      {message: '表の描画とセルへのフォーカス'});
    return sheet;
  };
  return env;
}

export class ExcelPage {
  constructor(page, html) {
    this.page = page;
    this.html = html;
  }

  evaluate(fn, ...args) { return this.page.evaluate(fn, ...args); }
  waitFor(fn, options, ...args) { return this.page.waitFor(fn, options, ...args); }
  press(combo) { return this.page.press(combo); }
  type(text) { return this.page.type(text); }

  /** Centre of the element matched by selector, scrolled into view. Fails when another element covers it. */
  async point(selector, {hitCheck = true} = {}) {
    const result = await this.evaluate((selector, hitCheck) => {
      const el = document.querySelector(selector);
      if (!el) return {error: '要素がありません: ' + selector};
      el.scrollIntoView({block: 'nearest', inline: 'nearest'});
      // scrollIntoView ignores the sticky column and row headers; move a cell out from under them.
      const viewport = document.getElementById('viewport');
      if (viewport.contains(el) && !el.closest('thead, .row-head')) {
        // The corner header stays at the top left, so its bottom and right edges are where the headers end.
        const corner = viewport.querySelector('#grid thead th').getBoundingClientRect();
        let box = el.getBoundingClientRect();
        if (box.top < corner.bottom) viewport.scrollTop -= corner.bottom - box.top;
        box = el.getBoundingClientRect();
        if (box.left < corner.right) viewport.scrollLeft -= corner.right - box.left;
      }
      const r = el.getBoundingClientRect(), x = r.left + r.width / 2, y = r.top + r.height / 2;
      if (hitCheck) {
        const hit = document.elementFromPoint(x, y);
        if (!hit || !(el === hit || el.contains(hit))) return {error: selector + ' の位置に別の要素があります: ' + (hit ? hit.outerHTML.slice(0, 120) : 'なし')};
      }
      return {x, y};
    }, selector, hitCheck);
    if (result.error) throw new Error(result.error);
    return result;
  }

  /** The element id of a cell; inside a merged area this is its top-left cell. */
  async cellSelector(addr) {
    const id = await this.evaluate(a => {
      if (document.getElementById('cell-' + a)) return 'cell-' + a;
      const C = window.ExcelHTMLCore, [r, c] = C.coord(a);
      for (const td of document.querySelectorAll('#grid td')) {
        const r1 = +td.dataset.r, c1 = +td.dataset.c;
        if (r >= r1 && r < r1 + td.rowSpan && c >= c1 && c < c1 + td.colSpan) return td.id;
      }
      return null;
    }, addr);
    if (!id) throw new Error('セルがありません: ' + addr);
    return '#' + id;
  }

  async click(selector, {shift = false, ctrl = false, button = 'left', hitCheck = true} = {}) {
    const point = await this.point(selector, {hitCheck});
    await this.page.click(point, {button, modifiers: (shift ? 8 : 0) | (ctrl ? 2 : 0)});
  }
  async clickCell(addr, options) { await this.click(await this.cellSelector(addr), options); }
  async dblclickCell(addr) { await this.page.dblclick(await this.point(await this.cellSelector(addr))); }
  async rightClickCell(addr) { await this.click(await this.cellSelector(addr), {button: 'right'}); }

  /** Selects a cell or range by typing it into the name box, then gives the focus back to the sheet. */
  async select(range) {
    await this.evaluate(() => { const box = document.getElementById('name-box'); box.focus(); box.select(); });
    await this.type(range);
    await this.press('Enter');
  }

  /** Clicks addr, types text (replacing the content, as in Excel) and confirms with Enter. */
  async enter(addr, text) {
    await this.clickCell(addr);
    await this.type(text);
    await this.press('Enter');
  }

  /** What the cell shows and how it is drawn. null for a cell hidden inside a merged area. */
  cell(addr) {
    return this.evaluate(a => {
      const td = document.getElementById('cell-' + a);
      if (!td) return null;
      const span = td.querySelector('.cell-text'), cs = getComputedStyle(td), r = td.getBoundingClientRect(), sr = span.getBoundingClientRect();
      return {
        text: span.textContent, title: td.title, colSpan: td.colSpan, rowSpan: td.rowSpan, classes: [...td.classList], link: td.dataset.link || null,
        rect: {left: r.left, top: r.top, width: r.width, height: r.height}, textRect: {left: sr.left, width: sr.width},
        style: {fontFamily: cs.fontFamily, fontSize: cs.fontSize, fontWeight: cs.fontWeight, fontStyle: cs.fontStyle, textDecorationLine: cs.textDecorationLine,
          color: cs.color, backgroundColor: cs.backgroundColor, textAlign: cs.textAlign, verticalAlign: cs.verticalAlign, whiteSpace: cs.whiteSpace,
          borderTop: cs.borderTopWidth + ' ' + cs.borderTopStyle + ' ' + cs.borderTopColor, borderRight: cs.borderRightWidth + ' ' + cs.borderRightStyle + ' ' + cs.borderRightColor,
          borderBottom: cs.borderBottomWidth + ' ' + cs.borderBottomStyle + ' ' + cs.borderBottomColor, borderLeft: cs.borderLeftWidth + ' ' + cs.borderLeftStyle + ' ' + cs.borderLeftColor,
          position: cs.position, zIndex: cs.zIndex},
        inline: {borderLeft: td.style.borderLeft, borderTop: td.style.borderTop, borderRight: td.style.borderRight, borderBottom: td.style.borderBottom},
      };
    }, addr);
  }
  async text(addr) { return (await this.cell(addr))?.text ?? null; }
  value(id) { return this.evaluate(id => document.getElementById(id).value, id); }
  textOf(selector) { return this.evaluate(s => document.querySelector(s)?.textContent ?? null, selector); }
  nameBox() { return this.value('name-box'); }
  formula() { return this.value('formula'); }
  status() { return this.textOf('#status'); }
  stats() { return this.textOf('#stats'); }
  selectionSize() { return this.textOf('#selection-size'); }
  activeElement() { return this.evaluate(() => document.activeElement?.id || document.activeElement?.tagName); }
  /** Address of the active cell (the white cell inside the selection). */
  activeCell() { return this.evaluate(() => document.querySelector('#grid td.active-cell')?.id.slice('cell-'.length) ?? null); }
  /** How far the green selection frame is from covering exactly the cells from..to (largest gap in px). */
  selectionFrameGap(from, to) {
    return this.evaluate((from, to) => {
      const frame = document.getElementById('selection').getBoundingClientRect();
      const a = document.getElementById('cell-' + from).getBoundingClientRect(), b = document.getElementById('cell-' + to).getBoundingClientRect();
      return Math.max(Math.abs(frame.left - a.left), Math.abs(frame.top - a.top), Math.abs(frame.right - b.right), Math.abs(frame.bottom - b.bottom));
    }, from, to);
  }
  /** Whether the cell editor is showing (typing goes into a cell). */
  editing() { return this.evaluate(() => document.getElementById('cell-editor').style.opacity === '1'); }

  /** Clicks a ribbon button by its data-action, first opening the ribbon tab that holds it. */
  async action(name) {
    const tab = await this.evaluate(n => document.querySelector(`.ribbon [data-action="${n}"]`)?.closest('.ribbon').id.slice('ribbon-'.length) ?? null, name);
    if (!tab) throw new Error('リボンにボタンがありません: ' + name);
    if (await this.evaluate(t => document.getElementById('ribbon-' + t).hidden, tab)) await this.click(`[data-tab="${tab}"]`);
    await this.click(`#ribbon-${tab} [data-action="${name}"]`);
  }

  /** Right-clicks addr and chooses the context-menu item with that data-action. */
  async contextAction(addr, name) {
    await this.rightClickCell(addr);
    await this.waitFor(() => !document.getElementById('context-menu').hidden, {message: '右クリックメニューの表示'});
    await this.click(`#context-menu [data-action="${name}"]`);
  }

  /** Chooses a value in a ribbon select or input (font, size, colours, alignment, number format, border...). */
  setControl(id, value) {
    return this.evaluate((id, value) => {
      const el = document.getElementById(id);
      el.value = String(value);
      el.dispatchEvent(new Event('input', {bubbles: true}));
      el.dispatchEvent(new Event('change', {bubbles: true}));
    }, id, value);
  }

  /** The open dialog as {id, title, body, input}, or null. */
  dialog() {
    return this.evaluate(() => {
      const d = document.querySelector('dialog[open]');
      if (!d) return null;
      const input = d.querySelector('#message-input');
      return {id: d.id, title: d.querySelector('h2')?.textContent ?? '', body: d.querySelector('p')?.textContent ?? '', input: input && !input.hidden ? input.value : null};
    });
  }
  waitDialog(id) {
    return this.waitFor(id => { const d = document.querySelector('dialog[open]'); return d && (!id || d.id === id) ? d.id : null; }, {message: 'ダイアログ ' + (id || '') + ' の表示'}, id);
  }
  /** Answers the message dialog: types input (when given) and presses OK ('ok') or キャンセル ('cancel'). */
  async answer(button = 'ok', input) {
    await this.waitDialog('message-dialog');
    if (input !== undefined) {
      await this.evaluate(() => { const box = document.getElementById('message-input'); box.focus(); box.select(); });
      await this.type(input);
    }
    // The page acts on the answer in the dialog's close event, which comes after the dialog has closed.
    // This listener is added after the page's own, so the page has finished by the time it fires. The same dialog
    // element may open again right away (an error message), so success is judged by this event, not by "open".
    await this.evaluate(() => {
      window.__testDialogClosed = false;
      document.getElementById('message-dialog').addEventListener('close', () => setTimeout(() => { window.__testDialogClosed = true; }), {once: true});
    });
    const selector = `#message-dialog button[value="${button}"]`;
    for (let attempt = 1; ; attempt++) {
      // A dialog that has just opened may still move while it lays out; press the button where it has come to rest.
      await this.evaluate(async s => {
        const el = document.querySelector(s), frame = () => new Promise(r => { requestAnimationFrame(() => r()); setTimeout(r, 50); });
        let last = '';
        for (let i = 0; i < 10; i++) {
          await frame();
          const r = el.getBoundingClientRect(), now = [r.left, r.top, r.width, r.height].join();
          if (now === last) return;
          last = now;
        }
      }, selector);
      await this.click(selector);
      try {
        await this.waitFor(() => window.__testDialogClosed, {timeout: 2000, message: 'ダイアログの答えの処理'});
        return;
      } catch (error) {
        if (attempt === 3) {
          const detail = await this.evaluate(() => { const d = document.getElementById('message-dialog'); return {open: d.open, visibility: document.visibilityState, focus: document.hasFocus()}; });
          throw new Error(error.message + ' ' + JSON.stringify(detail));
        }
      }
    }
  }

  /** The state text of 「AIに渡す」, parsed: {raw, state, book, sheet(id)}. */
  async state() {
    const raw = await this.evaluate(() => {
      document.getElementById('handoff').click();
      const text = document.getElementById('handoff-text').value;
      document.getElementById('handoff-dialog').close();
      return text;
    });
    const newline = raw.indexOf('\n');
    if (raw.slice(0, newline) !== 'EXCEL_HTML_STATE_V1') throw new Error('状態テキストの先頭が EXCEL_HTML_STATE_V1 ではありません');
    const state = JSON.parse(raw.slice(newline + 1));
    const book = state.workbook;
    return {raw, state, book, sheet: (id = book.activeSheet) => book.sheets.find(s => s.id === id)};
  }
  /** The cells of one sheet in the state text. */
  async cells(sheetId) { return (await this.state()).sheet(sheetId).cells; }

  /** Sends the copy event the browser fires for Ctrl+C and returns what the page put in it. */
  copy() { return this.clipboardEvent('copy'); }
  /** Sends the cut event the browser fires for Ctrl+X. */
  cut() { return this.clipboardEvent('cut'); }
  clipboardEvent(type) {
    return this.evaluate(type => {
      const data = new DataTransfer(), target = document.activeElement || document.body;
      const event = new ClipboardEvent(type, {clipboardData: data, bubbles: true, cancelable: true, composed: true});
      target.dispatchEvent(event);
      return {prevented: event.defaultPrevented, html: data.getData('text/html'), text: data.getData('text/plain'), types: [...data.types], target: target.id};
    }, type);
  }
  /** Sends the paste event the browser fires for Ctrl+V, carrying html and/or text. */
  paste({html = null, text = null} = {}) {
    return this.evaluate((html, text) => {
      const data = new DataTransfer(), target = document.activeElement || document.body;
      if (html !== null) data.setData('text/html', html);
      if (text !== null) data.setData('text/plain', text);
      const event = new ClipboardEvent('paste', {clipboardData: data, bubbles: true, cancelable: true, composed: true});
      target.dispatchEvent(event);
      return {prevented: event.defaultPrevented, target: target.id};
    }, html, text);
  }
  /** Pastes a PNG picture of width x height (drawn in the page), optionally together with html / text. */
  pastePicture({width = 120, height = 60, color = '#ff0000', html = null, text = null} = {}) {
    return this.evaluate(async (width, height, color, html, text) => {
      const canvas = document.createElement('canvas');
      canvas.width = width; canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = color; ctx.fillRect(0, 0, width, height);
      const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
      const data = new DataTransfer(), target = document.activeElement || document.body;
      data.items.add(new File([blob], 'picture.png', {type: 'image/png'}));
      if (html !== null) data.setData('text/html', html);
      if (text !== null) data.setData('text/plain', text);
      const event = new ClipboardEvent('paste', {clipboardData: data, bubbles: true, cancelable: true, composed: true});
      target.dispatchEvent(event);
      return {prevented: event.defaultPrevented, target: target.id};
    }, width, height, color, html, text);
  }

  /** What the in-page clipboard holds: [{type, size, text}]. */
  clipboard() { return this.evaluate(() => window.__testClipboard.dump()); }
  setClipboard(entries) { return this.evaluate(entries => window.__testClipboard.set(entries), entries); }

  /** Ends the test's use of the tab: fails when the page reported an error, and hands the tab back for reuse. */
  release(idle) {
    const errors = this.page.errors.splice(0);
    if (!this.page.closed) idle.push(this.page);
    if (errors.length) throw new Error('ページでエラーが発生しました:\n' + errors.join('\n'));
  }
}
