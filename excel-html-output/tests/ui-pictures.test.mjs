// 画像：貼り付け、移動、大きさの変更、削除と元に戻す、画像だけのコピー、表と一緒のときの扱い、重なり順と固定。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {fixture} from './lib/workbook.mjs';

const env = useBrowser();
const book = fixture('book');
const PNG_1PX = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const pictures = x => x.evaluate(() => [...document.querySelectorAll('#pictures .picture')].map(el => {
  const r = el.getBoundingClientRect();
  return {id: el.dataset.picture, picked: el.classList.contains('picked'), handle: !!el.querySelector('.picture-handle'), zIndex: el.style.zIndex,
    left: r.left, top: r.top, width: r.width, height: r.height};
}));

async function pastePicture(x, addr, options) {
  await x.clickCell(addr);
  await x.pastePicture(options);
  await x.waitFor(() => document.querySelector('#pictures .picture'), {message: '画像の表示'});
  return (await pictures(x))[0];
}

/** The cell under the top-left corner of the picture, and the offset from that cell's corner. */
const anchorOnScreen = (x, id) => x.evaluate(id => {
  const pic = document.querySelector(`[data-picture="${id}"]`).getBoundingClientRect();
  const td = [...document.querySelectorAll('#grid td')].find(td => { const r = td.getBoundingClientRect(); return pic.left >= r.left && pic.left < r.right && pic.top >= r.top && pic.top < r.bottom; });
  const r = td.getBoundingClientRect();
  return {r: +td.dataset.r, c: +td.dataset.c, x: Math.round(pic.left - r.left), y: Math.round(pic.top - r.top)};
}, id);

test('画像だけをコピーして Ctrl+V で貼ると、選んだセルの位置に、選ばれた状態で置かれる', async t => {
  const x = await env.open(t, book);
  const pic = await pastePicture(x, 'C5', {width: 120, height: 60});
  const c5 = await x.cell('C5');
  assert.deepEqual([pic.width, pic.height], [120, 60], '画面に見えていた大きさ（縮尺1倍）');
  assert.ok(Math.abs(pic.left - c5.rect.left) <= 1 && Math.abs(pic.top - c5.rect.top) <= 1, 'C5 の左上に置く');
  assert.deepEqual([pic.picked, pic.handle], [true, true], '貼った画像は選ばれ、大きさを変えるつまみが出る');
  assert.match(await x.status(), /画像を貼り付けました/);
  const {state, sheet} = await x.state();
  const image = sheet('main').images[0];
  assert.deepEqual({r: image.r, c: image.c, x: image.x, y: image.y, w: image.w, h: image.h}, {r: 4, c: 2, x: 0, y: 0, w: 120, h: 60});
  assert.match(state.workbook.assets[image.asset], /^data:image\/png;base64,/, '状態テキストには画像の中身も入る');
});

test('画像をドラッグすると移動し、移動先のセルを基準に位置を持ち直す', async t => {
  const x = await env.open(t, book);
  const pic = await pastePicture(x, 'C5');
  const center = {x: pic.left + pic.width / 2, y: pic.top + pic.height / 2};
  await x.page.drag(center, {x: center.x + 200, y: center.y + 50});
  const moved = (await pictures(x))[0];
  assert.ok(Math.abs(moved.left - pic.left - 200) <= 1 && Math.abs(moved.top - pic.top - 50) <= 1, '画面上で 200px 右、50px 下へ動く');
  const image = (await x.state()).sheet('main').images[0];
  assert.deepEqual({r: image.r, c: image.c, x: image.x, y: image.y}, await anchorOnScreen(x, image.id));
  assert.notDeepEqual([image.r, image.c], [4, 2], '基準のセルが C5 から変わる');
});

test('右下のつまみのドラッグで、縦横比を保ったまま大きさを変える', async t => {
  const x = await env.open(t, book);
  await pastePicture(x, 'C5', {width: 120, height: 60});
  const handle = await x.point('.picture-handle');
  await x.page.drag(handle, {x: handle.x + 60, y: handle.y + 5});
  const image = (await x.state()).sheet('main').images[0];
  assert.deepEqual([image.w, image.h], [180, 90]);
});

test('選んだ画像は Delete で削除し、元に戻すと画像も戻る', async t => {
  const x = await env.open(t, book);
  await pastePicture(x, 'C5');
  await x.press('Delete');
  assert.deepEqual(await pictures(x), []);
  assert.equal(await x.status(), '画像を削除しました');
  const after = await x.state();
  assert.equal(after.sheet('main').images, undefined);
  assert.equal(after.state.workbook.assets, undefined, '使われていない画像の中身は、状態テキストに入れない');
  assert.equal(await x.text('C5'), '2', '画像の下のセルの値（C5 の 2）は消さない');
  await x.press('Control+z');
  assert.equal((await pictures(x)).length, 1);
});

test('選んだ画像の Ctrl+C は画像だけをコピーし、Ctrl+X は画像をコピーしてから削除する', async t => {
  const x = await env.open(t, book);
  await pastePicture(x, 'C5');
  const copied = await x.copy();
  assert.equal(copied.prevented, true);
  assert.equal(copied.html, '', 'セルの範囲は渡さない');
  await x.waitFor(() => window.__testClipboard.writes > 0, {message: '画像のコピー'});
  const stored = await x.clipboard();
  assert.equal(stored.length, 1);
  assert.equal(stored[0].type, 'image/png');
  assert.ok(stored[0].size > 0);
  assert.match(await x.status(), /画像をコピーしました/);
  await x.cut();
  await x.waitFor(() => !document.querySelector('#pictures .picture'), {message: '切り取った画像の削除'});
  assert.equal(await x.status(), '画像を切り取りました');
});

test('矢印キーや Esc で画像の選択を外し、セルの選択に戻る', async t => {
  const x = await env.open(t, book);
  await pastePicture(x, 'C5');
  assert.equal(await x.evaluate(() => document.getElementById('selection').style.visibility), 'hidden', '画像を選んでいる間はセルの選択枠を隠す');
  await x.press('Escape');
  assert.equal((await pictures(x))[0].picked, false);
  assert.equal(await x.evaluate(() => document.getElementById('selection').style.visibility), '');
  await x.click('#pictures .picture');
  assert.equal((await pictures(x))[0].picked, true, 'クリックで選び直せる');
  await x.press('ArrowDown');
  assert.equal((await pictures(x))[0].picked, false);
  assert.equal(await x.nameBox(), 'C6');
});

test('クリップボードに表や文字があれば、画像が一緒でもセルとして貼る（Excelやウェブの表のコピー）', async t => {
  const x = await env.open(t, book);
  await x.clickCell('G4');
  await x.pastePicture({html: '<table><tr><td>表の値</td></tr></table>', text: '表の値'});
  await x.waitFor(() => document.getElementById('cell-G4').textContent === '表の値', {message: 'セルとしての貼り付け'});
  await x.clickCell('G6');
  await x.pastePicture({text: '文字だけ'});
  await x.waitFor(() => document.getElementById('cell-G6').textContent === '文字だけ', {message: '文字の貼り付け'});
  assert.deepEqual(await pictures(x), [], '画像は置かない');
});

test('行を挿入すると、画像も基準のセルと一緒に下へずれる', async t => {
  const x = await env.open(t, book);
  await pastePicture(x, 'C5');
  await x.press('Escape');
  await x.clickCell('A3');
  await x.action('insert-row');
  assert.equal((await x.state()).sheet('main').images[0].r, 5);
});

test('画像の重なり順は、固定しない位置で3、固定した行か列で5、両方で7になり、固定した位置の画像はスクロールしても動かない', async t => {
  const large = fixture('book');
  Object.assign(large.sheets[0], {rows: 120, cols: 40, freeze: {rows: 3, cols: 3}, merges: [], images: [
    {id: 'both', asset: 'img', r: 0, c: 1, x: 2, y: 2, w: 40, h: 20},
    {id: 'row', asset: 'img', r: 1, c: 6, x: 0, y: 0, w: 40, h: 20},
    {id: 'free', asset: 'img', r: 10, c: 6, x: 0, y: 0, w: 40, h: 20},
  ]});
  delete large.sheets[0].cells.A11;
  large.assets = {img: PNG_1PX};
  const x = await env.open(t, large);
  const byId = async () => Object.fromEntries((await pictures(x)).map(p => [p.id, p]));
  const before = await byId();
  assert.deepEqual([before.both.zIndex, before.row.zIndex, before.free.zIndex], ['7', '5', '3']);
  await x.evaluate(() => { const v = document.getElementById('viewport'); v.scrollTop = 150; v.scrollLeft = 100; v.dispatchEvent(new Event('scroll')); });
  await x.waitFor(() => document.getElementById('viewport').scrollTop === 150, {message: 'スクロール'});
  const after = await byId();
  assert.deepEqual([after.both.left, after.both.top], [before.both.left, before.both.top], '行と列を固定した位置の画像は動かない');
  assert.equal(after.row.top, before.row.top, '固定した行の画像は上下に動かない');
  assert.equal(after.row.left, before.row.left - 100, '固定した行の画像も、左右にはスクロールに合わせて動く');
  assert.deepEqual([after.free.left, after.free.top], [before.free.left - 100, before.free.top - 150]);
});
