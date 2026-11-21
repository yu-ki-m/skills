// 画面のスクロールと見える位置：キー操作で移った先のセルが、列見出しや行番号に隠れないこと（矢印・PageUp・Ctrl+Home・
// 名前ボックス・範囲の中の Enter・Shift+矢印、拡大した表）。枠を固定した表のテストは ui-scroll-frozen にある（ただし、名前
// ボックスのテストの最後の確認と、拡大した表のテストは、枠を固定した表でここで確かめる）。シートの追加・見出しと選択の枠の
// 重なりは ui-viewport に、拡大・縮小は ui-zoom に、ドラッグは ui-drag にある（共通の部品は lib/view.mjs）。
import assert from 'node:assert/strict';
import {test} from 'node:test';
import {useBrowser} from './lib/app.mjs';
import {hidden, press} from './lib/view.mjs';
import {blankBook} from './lib/workbook.mjs';

const env = useBrowser();

test('スクロールした表でキーボードで上や左へ移っても、アクティブセルは列見出しや行番号に隠れずに見える', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40}}));
  await x.select('A61');
  await x.press('Control+Home');
  assert.equal(await hidden(x), null, '下へ移ってから Ctrl+Home で A1 に戻ったとき（1行目が列見出しの下に入らない）');
  await x.select('A60');
  await press(x, 'ArrowUp', 40);
  assert.equal(await hidden(x), null, '↑で画面の上の外へ出たとき');
  await x.select('A100');
  await press(x, 'PageUp', 3);
  assert.equal(await hidden(x), null, 'PageUp を続けて押したとき');
  await x.select('A5');
  await press(x, 'ArrowRight', 35);
  await press(x, 'ArrowLeft', 34);
  assert.equal(await x.nameBox(), 'B5');
  assert.equal(await hidden(x), null, '右へスクロールした後に←で戻ったとき（左の端が行番号の下に入らない）');
  await x.select('A1');
  await press(x, 'ArrowDown', 60);
  assert.equal(await hidden(x), null, '↓で画面の下の外へ出たとき');
  await x.select('A1');
  await press(x, 'ArrowRight', 30);
  assert.equal(await hidden(x), null, '→で画面の右の外へ出たとき');
});

test('範囲の中を Enter で巡って画面の上の外のセルに移っても、隠れずに見える', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 12}}));
  const active = () => x.evaluate(() => document.querySelector('#grid td.active-cell').id);
  await x.select('B1:B90');
  await x.press('Shift+Enter');
  assert.equal(await active(), 'cell-B90', 'Shift+Enter で、範囲の最後のセルへ巡る');
  assert.notEqual(await hidden(x, 'B1'), null, '前提: このとき B1 は画面の上の外にある');
  await x.press('Enter');
  assert.equal(await x.nameBox(), 'B1:B90');
  assert.equal(await active(), 'cell-B1', 'Enter で、範囲の最初のセルへ巡る');
  assert.equal(await hidden(x), null);
});

test('Shift+矢印で範囲を上や左へ広げても、動いている端のセルは列見出しや行番号に隠れずに見える', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40}}));
  await x.select('A60');
  await press(x, 'Shift+ArrowUp', 40);
  assert.equal(await x.nameBox(), 'A20:A60');
  assert.equal(await hidden(x, 'A20'), null, 'Shift+↑で画面の上の外へ広げたとき');
  await x.select('A5');
  await press(x, 'ArrowRight', 35);
  await press(x, 'Shift+ArrowLeft', 34);
  assert.equal(await x.nameBox(), 'B5:AJ5');
  assert.equal(await hidden(x, 'B5'), null, 'Shift+←で画面の左の外へ広げたとき');
});

// Excel shows the active cell (the top-left one) of a range chosen in the name box: the whole range when it fits.
test('名前ボックスで今の表示より上や左の範囲を選んでも、アクティブセル（範囲の左上）が見える', async t => {
  const x = await env.open(t, blankBook({sheet: {rows: 120, cols: 40}}));
  await x.select('A100');
  await x.select('A5:B6');
  assert.equal(await hidden(x), null, '上の2行の範囲: A5 が列見出しの下に入らない');
  assert.equal(await hidden(x, 'B6'), null, '画面に収まる範囲は、右下のセルも見える');
  await x.select('A100');
  await x.select('A1:B5');
  assert.equal(await hidden(x), null, '上の5行の範囲: A1 が画面の上の外に出ない');
  await x.select('AD5');
  await x.select('E5:F6');
  assert.equal(await hidden(x), null, '左の範囲: E5 が画面の左の外に出ない');
  await x.select('A1');
  await x.select('B1:B90');
  assert.equal(await hidden(x), null, '画面に収まらない範囲では、左上のセルが見える');
  const frozen = await env.open(t, blankBook({sheet: {rows: 120, cols: 40, freeze: {rows: 2, cols: 1}}}));
  await frozen.select('B100');
  await frozen.select('B10:C11');
  assert.equal(await hidden(frozen), null, '枠を固定した表: B10 が固定した行の下に入らない');
});

test('拡大した表でも、上や左へ移ったセルは見出しに隠れない', async t => {
  const x = await env.open(t, blankBook({view: {zoom: 150}, sheet: {rows: 120, cols: 40, freeze: {rows: 1, cols: 1}}}));
  await x.select('B60');
  await press(x, 'ArrowUp', 40);
  assert.equal(await hidden(x), null, '↑で戻ったとき');
  await x.select('B5');
  await press(x, 'ArrowRight', 30);
  await press(x, 'ArrowLeft', 28);
  assert.equal(await hidden(x), null, '←で戻ったとき');
});
