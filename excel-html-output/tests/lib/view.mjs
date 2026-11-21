// Helpers for the tests of what the sheet shows where (tests/ui-scroll.test.mjs, tests/ui-scroll-frozen.test.mjs,
// tests/ui-viewport.test.mjs, tests/ui-zoom.test.mjs, tests/ui-drag.test.mjs): whether a cell is hidden under the headers
// or out of view, the scroll offsets of the sheet, a key pressed several times, the box of an element and its middle, and
// a scroll made as the scroll bars make it.

// The column headers and the row numbers stay over the sheet's top and left edges (and the frozen rows and columns
// below and beside them), so a cell scrolled in from above or the left must land clear of them, not under them.
// Whether a cell (the active cell unless an address is given) is seen is checked at its corners and centre (the fill
// handle sits on the bottom-right one).
export const hidden = (x, address) => x.evaluate(a => {
  const td = a ? document.getElementById('cell-' + a) : document.querySelector('#grid td.active-cell'), r = td.getBoundingClientRect();
  for (const [name, px, py] of [['左上', r.left + 4, r.top + 4], ['右上', r.right - 4, r.top + 4], ['左下', r.left + 4, r.bottom - 4], ['中央', (r.left + r.right) / 2, (r.top + r.bottom) / 2]]) {
    const hit = document.elementFromPoint(px, py);
    if (!(td === hit || td.contains(hit))) return td.id + ' の' + name + 'が ' + (hit ? hit.tagName + (hit.className ? '.' + hit.className : '') : '画面の外') + ' の下にある';
  }
  return null;
}, address);
export const scrolled = x => x.evaluate(() => { const v = document.getElementById('viewport'); return {top: v.scrollTop, left: v.scrollLeft}; });
export const press = async (x, key, times) => { for (let i = 0; i < times; i++) await x.press(key); };
export const box = (x, selector) => x.evaluate(s => { const r = document.querySelector(s).getBoundingClientRect(); return {left: r.left, top: r.top, right: r.right, bottom: r.bottom, width: r.width}; }, selector);
export const middle = b => ({x: (b.left + b.right) / 2, y: (b.top + b.bottom) / 2});
// Scrolls the sheet by dy, dx as the scroll bars and the mouse wheel do, and returns once the page has handled the scroll
// (this listener comes after the page's own), or after a second when the sheet could not scroll.
export const scrollBy = (x, dy, dx = 0) => x.evaluate((dy, dx) => new Promise(resolve => {
  const v = document.getElementById('viewport');
  v.addEventListener('scroll', () => resolve(), {once: true});
  setTimeout(resolve, 1000);
  v.scrollTop += dy; v.scrollLeft += dx;
}), dy, dx);
