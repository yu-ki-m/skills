// Drives Chrome or Edge without a window over the Chrome DevTools Protocol (CDP).
// Only Node.js 22+ built-ins are used (global WebSocket, child_process), so no npm package is needed.
//
// Each page is a separate tab of one browser process; input goes through CDP's Input domain, which the page
// receives as trusted keyboard, mouse and IME events, the same way as from a person.
import {spawn} from 'node:child_process';
import {existsSync, readFileSync} from 'node:fs';
import {join} from 'node:path';
import {setTimeout as sleep} from 'node:timers/promises';
import {makeTempDir, removeTempDir} from './tempdirs.mjs';

/** Path of the browser to test with: EXCEL_HTML_TEST_BROWSER, else Chrome, else Edge. null when none is installed. */
export function findBrowser() {
  const chosen = process.env.EXCEL_HTML_TEST_BROWSER;
  if (chosen) return existsSync(chosen) ? chosen : null;
  const candidates = [];
  if (process.platform === 'win32') {
    const roots = [process.env.PROGRAMFILES, process.env['PROGRAMFILES(X86)'], process.env.LOCALAPPDATA].filter(Boolean);
    for (const root of roots) candidates.push(join(root, 'Google', 'Chrome', 'Application', 'chrome.exe'));
    for (const root of roots) candidates.push(join(root, 'Microsoft', 'Edge', 'Application', 'msedge.exe'));
  } else if (process.platform === 'darwin') {
    candidates.push('/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
      '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge', '/Applications/Chromium.app/Contents/MacOS/Chromium');
  } else {
    candidates.push('/usr/bin/google-chrome', '/usr/bin/google-chrome-stable', '/usr/bin/chromium', '/usr/bin/chromium-browser',
      '/usr/bin/microsoft-edge', '/snap/bin/chromium');
  }
  return candidates.find(path => existsSync(path)) || null;
}

// One WebSocket carries the browser session and every page session ("flatten" mode); messages are routed by sessionId.
class Connection {
  constructor(socket) {
    this.socket = socket;
    this.lastId = 0;
    this.pending = new Map();
    this.handlers = new Map();
    this.closed = false;
    socket.onmessage = message => {
      const data = JSON.parse(message.data);
      if (data.id !== undefined) {
        const call = this.pending.get(data.id);
        if (!call) return;
        this.pending.delete(data.id);
        if (data.error) call.reject(new Error(`${call.method}: ${data.error.message}${data.error.data ? ' (' + data.error.data + ')' : ''}`));
        else call.resolve(data.result);
        return;
      }
      for (const handler of this.handlers.get(data.sessionId || '') || []) handler(data.method, data.params || {});
    };
    socket.onclose = () => {
      this.closed = true;
      for (const call of this.pending.values()) call.reject(new Error(`${call.method}: ブラウザーとの接続が切れました`));
      this.pending.clear();
    };
  }

  // Every request has a time limit, so that a browser which stops answering fails the test instead of hanging it.
  send(method, params = {}, sessionId, timeout = 30000) {
    if (this.closed) return Promise.reject(new Error(`${method}: ブラウザーとの接続が切れています`));
    const id = ++this.lastId;
    const message = {id, method, params};
    if (sessionId) message.sessionId = sessionId;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        if (this.pending.delete(id)) reject(new Error(`${method}: ブラウザーが ${timeout / 1000} 秒以内に応答しませんでした`));
      }, timeout);
      timer.unref();
      const done = fn => value => { clearTimeout(timer); fn(value); };
      this.pending.set(id, {resolve: done(resolve), reject: done(reject), method});
      this.socket.send(JSON.stringify(message));
    });
  }

  listen(sessionId, handler) {
    const key = sessionId || '';
    if (!this.handlers.has(key)) this.handlers.set(key, new Set());
    this.handlers.get(key).add(handler);
    return () => this.handlers.get(key)?.delete(handler);
  }
}

/** Starts a browser without a window. Close it with browser.close(), normally in an after() hook. */
export async function launch({executable = findBrowser(), headed = process.env.EXCEL_HTML_TEST_HEADED === '1'} = {}) {
  if (!executable) throw new Error('Chrome または Edge が見つかりません。環境変数 EXCEL_HTML_TEST_BROWSER に実行ファイルのパスを指定できます。');
  const profile = makeTempDir('profile-');
  const args = [
    headed ? null : '--headless=new',
    '--remote-debugging-port=0', `--user-data-dir=${profile}`,
    '--no-first-run', '--no-default-browser-check', '--disable-extensions', '--disable-sync', '--disable-default-apps',
    // No traffic of the browser's own: updates, crash reports, field trials, and the models that Optimization Guide
    // downloads in the background (they arrived as tens of MB of .crx files while the tests ran).
    '--disable-background-networking', '--disable-component-update', '--disable-component-extensions-with-background-pages',
    '--disable-field-trial-config', '--disable-client-side-phishing-detection', '--disable-domain-reliability', '--disable-breakpad',
    '--metrics-recording-only', '--no-service-autorun', '--password-store=basic', '--disable-search-engine-choice-screen',
    '--disable-features=Translate,OptimizationHints,OptimizationHintsFetching,OptimizationGuideModelDownloading,OptimizationTargetPrediction,MediaRouter,AutofillServerCommunication,CertificateTransparencyComponentUpdater',
    // Draw on the CPU: browsers running side by side then do not queue for one GPU (measured: about 20% faster).
    '--disable-gpu',
    // Fixed size and pixel ratio, so that positions and pasted picture sizes do not depend on the machine.
    // A smaller window has less to draw, which also makes the tests faster.
    '--window-size=1280,800', '--force-device-scale-factor=1', '--lang=ja-JP',
    'about:blank',
  ].filter(Boolean);
  const child = spawn(executable, args, {stdio: 'ignore'});
  // The browser must never keep the test process alive: it is killed when the process exits, whatever happened.
  child.unref();
  const killOnExit = () => { try { child.kill(); } catch {} };
  process.once('exit', killOnExit);
  let exited = null;
  child.once('exit', code => { exited = code ?? -1; });
  child.once('error', error => { exited = error; });

  try {
    const portFile = join(profile, 'DevToolsActivePort');
    const started = Date.now();
    let lines = [];
    while (lines.length < 2 || !lines[0]) {
      if (exited !== null) throw new Error(`ブラウザーが起動直後に終了しました（${executable}）: ${exited}`);
      if (Date.now() - started > 20000) throw new Error(`ブラウザーの起動を待ちきれませんでした（${executable}）`);
      try {
        lines = readFileSync(portFile, 'utf8').split(/\r?\n/);
      } catch (error) {
        // Not written yet, or still locked by the browser while it writes (EBUSY on Windows).
        if (!['ENOENT', 'EBUSY', 'EPERM', 'EACCES'].includes(error.code)) throw error;
      }
      if (lines.length < 2 || !lines[0]) await sleep(30);
    }
    const socket = new WebSocket(`ws://127.0.0.1:${lines[0]}${lines[1]}`);
    await new Promise((resolve, reject) => {
      socket.onopen = resolve;
      socket.onerror = () => reject(new Error('ブラウザーの操作用の接続（CDP）を開けませんでした'));
    });
    const connection = new Connection(socket);
    // Nothing is saved as a download, except while a test waits for one (Page.download).
    await connection.send('Browser.setDownloadBehavior', {behavior: 'deny', eventsEnabled: true});
    return new Browser(child, connection, profile, killOnExit);
  } catch (error) {
    killOnExit();
    process.removeListener('exit', killOnExit);
    await sleep(300);
    removeTempDir(profile);
    throw error;
  }
}

export class Browser {
  constructor(child, connection, profile, killOnExit) {
    this.child = child;
    this.connection = connection;
    this.profile = profile;
    this.killOnExit = killOnExit;
    this.pages = new Set();
    this.downloads = new Map();
    connection.listen('', (method, params) => {
      if (method === 'Browser.downloadWillBegin') this.downloads.set(params.guid, {...params, state: 'inProgress'});
      if (method === 'Browser.downloadProgress' && this.downloads.has(params.guid)) Object.assign(this.downloads.get(params.guid), {state: params.state});
    });
  }

  send(method, params, timeout) { return this.connection.send(method, params, undefined, timeout); }

  async newPage() {
    const {targetId} = await this.send('Target.createTarget', {url: 'about:blank'});
    const {sessionId} = await this.send('Target.attachToTarget', {targetId, flatten: true});
    const page = new Page(this, targetId, sessionId);
    await page.init();
    this.pages.add(page);
    return page;
  }

  async close() {
    for (const page of [...this.pages]) await page.close().catch(() => {});
    const exited = new Promise(resolve => {
      if (this.child.exitCode !== null) resolve();
      else this.child.once('exit', resolve);
    });
    await this.send('Browser.close', {}, 5000).catch(() => {});
    await Promise.race([exited, sleep(5000)]);
    this.killOnExit();
    process.removeListener('exit', this.killOnExit);
    removeTempDir(this.profile);
  }
}

// CDP modifier bits.
const MODIFIER = {Alt: 1, Control: 2, Meta: 4, Shift: 8};
const KEYS = {
  Enter: {code: 'Enter', keyCode: 13, text: '\r'}, Tab: {code: 'Tab', keyCode: 9}, Escape: {code: 'Escape', keyCode: 27},
  Backspace: {code: 'Backspace', keyCode: 8}, Delete: {code: 'Delete', keyCode: 46},
  Home: {code: 'Home', keyCode: 36}, End: {code: 'End', keyCode: 35}, PageUp: {code: 'PageUp', keyCode: 33}, PageDown: {code: 'PageDown', keyCode: 34},
  ArrowUp: {code: 'ArrowUp', keyCode: 38}, ArrowDown: {code: 'ArrowDown', keyCode: 40}, ArrowLeft: {code: 'ArrowLeft', keyCode: 37}, ArrowRight: {code: 'ArrowRight', keyCode: 39},
  F2: {code: 'F2', keyCode: 113}, F4: {code: 'F4', keyCode: 115},
  Shift: {code: 'ShiftLeft', keyCode: 16}, Control: {code: 'ControlLeft', keyCode: 17}, Alt: {code: 'AltLeft', keyCode: 18}, Meta: {code: 'MetaLeft', keyCode: 91},
};
function describeKey(key) {
  if (KEYS[key]) return {key, ...KEYS[key]};
  if ([...key].length !== 1) throw new Error(`未対応のキー: ${key}`);
  const upper = key.toUpperCase();
  const code = /^[a-z]$/i.test(key) ? 'Key' + upper : /^\d$/.test(key) ? 'Digit' + key : '';
  return {key, code, keyCode: code ? upper.charCodeAt(0) : 0, text: key};
}

export class Page {
  constructor(browser, targetId, sessionId) {
    this.browser = browser;
    this.targetId = targetId;
    this.sessionId = sessionId;
    this.errors = [];
    this.nativeDialogs = [];
    this.closed = false;
    this.stopListening = browser.connection.listen(sessionId, (method, params) => this.handle(method, params));
    this.waiters = new Set();
    this.mouse = {
      move: (x, y, options = {}) => this.mouseEvent('mouseMoved', x, y, options),
      down: (x, y, options = {}) => this.mouseEvent('mousePressed', x, y, {clickCount: 1, ...options}),
      up: (x, y, options = {}) => this.mouseEvent('mouseReleased', x, y, {clickCount: 1, ...options}),
    };
  }

  send(method, params) { return this.browser.connection.send(method, params, this.sessionId); }

  async init() {
    await this.send('Page.enable');
    await this.send('Runtime.enable');
    // Behave as the focused window, so that focus and :focus work even when several tabs exist.
    await this.send('Emulation.setFocusEmulationEnabled', {enabled: true});
  }

  handle(method, params) {
    if (method === 'Runtime.exceptionThrown') {
      const d = params.exceptionDetails;
      this.errors.push('例外: ' + (d.exception?.description || d.text));
    } else if (method === 'Runtime.consoleAPICalled' && (params.type === 'error' || params.type === 'assert')) {
      this.errors.push('console.' + params.type + ': ' + params.args.map(a => a.value ?? a.description ?? '').join(' '));
    } else if (method === 'Page.javascriptDialogOpening') {
      // The workbook never uses alert/confirm/prompt; only "leave page?" may appear when a dirty tab closes.
      this.nativeDialogs.push({type: params.type, message: params.message});
      if (params.type !== 'beforeunload') this.errors.push(`ブラウザー標準のダイアログ（${params.type}）が開きました: ${params.message}`);
      this.send('Page.handleJavaScriptDialog', {accept: true}).catch(() => {});
    }
    for (const waiter of [...this.waiters]) if (waiter.method === method) { this.waiters.delete(waiter); waiter.resolve(params); }
  }

  waitForEvent(method, timeout = 15000) {
    let waiter;
    const promise = new Promise((resolve, reject) => {
      waiter = {method, resolve};
      this.waiters.add(waiter);
      setTimeout(() => { if (this.waiters.delete(waiter)) reject(new Error(`${method} を待ちきれませんでした`)); }, timeout).unref();
    });
    return promise;
  }

  /** Runs source in every document of this tab before the page's own scripts. */
  addInitScript(source) { return this.send('Page.addScriptToEvaluateOnNewDocument', {source}); }

  async goto(url) {
    const loaded = this.waitForEvent('Page.loadEventFired');
    const result = await this.send('Page.navigate', {url});
    if (result.errorText) throw new Error(`${url} を開けませんでした: ${result.errorText}`);
    await loaded;
  }

  /** Evaluates fn(...args) in the page and returns the value (JSON-serialisable values only). */
  async evaluate(fn, ...args) {
    const expression = typeof fn === 'function' ? `(${fn})(...${JSON.stringify(args)})` : String(fn);
    const result = await this.send('Runtime.evaluate', {expression, awaitPromise: true, returnByValue: true, userGesture: true});
    if (result.exceptionDetails) {
      const d = result.exceptionDetails;
      throw new Error('ページ内の処理が失敗しました: ' + (d.exception?.description || d.text));
    }
    return result.result.value;
  }

  /** Polls fn(...args) in the page until it returns a truthy value, and returns that value. */
  async waitFor(fn, {timeout = 5000, message = '条件'} = {}, ...args) {
    const until = Date.now() + timeout;
    let last;
    for (;;) {
      last = await this.evaluate(fn, ...args);
      if (last) return last;
      if (Date.now() > until) throw new Error(`${message} を ${timeout}ms 待ちましたが、満たされませんでした（最後の値: ${JSON.stringify(last)}）`);
      await sleep(15);
    }
  }

  /** Presses a key combination such as 'Enter', 'Shift+Tab', 'Control+Z', 'Control+Shift+Z', 'Alt+Enter'. */
  async press(combo) {
    const parts = combo.split('+');
    const main = describeKey(parts.pop());
    let modifiers = 0;
    for (const name of parts) {
      if (!(name in MODIFIER)) throw new Error(`未対応の修飾キー: ${name}`);
      modifiers |= MODIFIER[name];
      await this.send('Input.dispatchKeyEvent', {type: 'rawKeyDown', key: name, code: KEYS[name].code, windowsVirtualKeyCode: KEYS[name].keyCode, modifiers});
    }
    // With Ctrl, Alt or Meta held, a key produces no character, as on a real keyboard.
    const text = modifiers & (MODIFIER.Alt | MODIFIER.Control | MODIFIER.Meta) ? undefined : main.text;
    const base = {key: main.key, code: main.code, windowsVirtualKeyCode: main.keyCode, nativeVirtualKeyCode: main.keyCode, modifiers};
    await this.send('Input.dispatchKeyEvent', {...base, type: text ? 'keyDown' : 'rawKeyDown', ...(text ? {text, unmodifiedText: text} : {})});
    await this.send('Input.dispatchKeyEvent', {...base, type: 'keyUp'});
    for (const name of parts.reverse()) {
      modifiers &= ~MODIFIER[name];
      await this.send('Input.dispatchKeyEvent', {type: 'keyUp', key: name, code: KEYS[name].code, windowsVirtualKeyCode: KEYS[name].keyCode, modifiers});
    }
  }

  /** Types text into the focused element, as committed input (also used for the text an IME commits). */
  type(text) { return this.send('Input.insertText', {text}); }

  /** Shows text as an unconfirmed IME composition (before the person presses Enter to confirm it). */
  compose(text) { return this.send('Input.imeSetComposition', {text, selectionStart: text.length, selectionEnd: text.length}); }

  mouseEvent(type, x, y, {button = 'left', clickCount = 0, modifiers = 0, buttons} = {}) {
    const pressed = buttons ?? (type === 'mousePressed' ? {left: 1, right: 2, middle: 4}[button] : 0);
    return this.send('Input.dispatchMouseEvent', {type, x, y, button: type === 'mouseMoved' && !pressed ? 'none' : button, buttons: pressed, clickCount, modifiers});
  }

  async click({x, y}, {button = 'left', modifiers = 0} = {}) {
    await this.mouse.move(x, y, {modifiers});
    await this.mouse.down(x, y, {button, modifiers});
    await this.mouse.up(x, y, {button, modifiers});
  }

  async dblclick({x, y}) {
    await this.mouse.move(x, y);
    await this.mouse.down(x, y, {clickCount: 1});
    await this.mouse.up(x, y, {clickCount: 1});
    await this.mouse.down(x, y, {clickCount: 2});
    await this.mouse.up(x, y, {clickCount: 2});
  }

  /** Presses the left button at from, moves in steps to to, and releases there. */
  async drag(from, to, {steps = 6, modifiers = 0} = {}) {
    await this.mouse.move(from.x, from.y, {modifiers});
    await this.mouse.down(from.x, from.y, {modifiers});
    for (let i = 1; i <= steps; i++) {
      const x = from.x + (to.x - from.x) * i / steps, y = from.y + (to.y - from.y) * i / steps;
      await this.mouse.move(x, y, {buttons: 1, button: 'left', modifiers});
    }
    await this.mouse.up(to.x, to.y, {modifiers});
  }

  /** Runs trigger() and returns {suggestedFilename, path} of the file the page downloads into dir. */
  async download(dir, trigger, timeout = 10000) {
    await this.browser.send('Browser.setDownloadBehavior', {behavior: 'allowAndName', downloadPath: dir, eventsEnabled: true});
    try {
      const before = new Set(this.browser.downloads.keys());
      await trigger();
      const until = Date.now() + timeout;
      for (;;) {
        const item = [...this.browser.downloads.values()].find(d => !before.has(d.guid) && d.frameId);
        if (item?.state === 'completed') return {suggestedFilename: item.suggestedFilename, path: join(dir, item.guid)};
        if (item?.state === 'canceled') throw new Error('ダウンロードが取り消されました');
        if (Date.now() > until) throw new Error('ダウンロードを待ちきれませんでした');
        await sleep(20);
      }
    } finally {
      await this.browser.send('Browser.setDownloadBehavior', {behavior: 'deny', eventsEnabled: true});
    }
  }

  async close() {
    if (this.closed) return;
    this.closed = true;
    this.browser.pages.delete(this);
    await this.browser.send('Target.closeTarget', {targetId: this.targetId}, 5000).catch(() => {});
    this.stopListening();
  }
}
