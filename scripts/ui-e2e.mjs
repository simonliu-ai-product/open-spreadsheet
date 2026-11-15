// End-to-end check of the grid UI in Chromium: starts `open-spreadsheet dev`
// on a fresh workspace, drives home, templates, editing, menus, sheets,
// formats, fill, freeze, merge, find, undo and xlsx round trip, and saves
// screenshots.
//
//   node scripts/ui-e2e.mjs <out dir>
import { spawn } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const out = path.resolve(process.argv[2] ?? 'e2e-out');
mkdirSync(out, { recursive: true });
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const bin = path.join(root, 'packages/core/bin.js');
const ws = path.join(mkdtempSync(path.join(tmpdir(), 'oss-e2e-')), 'book');
await new Promise((res, rej) => spawn(process.execPath, [bin, 'init', ws], { stdio: 'inherit' }).on('exit', (c) => (c ? rej(new Error('init')) : res())));
const PORT = 5395;
const server = spawn(process.execPath, [bin, 'dev', '--port', String(PORT)], { cwd: ws, stdio: ['ignore', 'pipe', 'inherit'], detached: true });
await new Promise((res) => server.stdout.on('data', (d) => String(d).includes('Local') && res()));
const URL = `http://localhost:${PORT}/`;

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok: Boolean(ok), detail });
  console.log(`${ok ? '✓' : '✗'} ${name}${detail ? ` — ${detail}` : ''}`);
};

const browser = await chromium.launch();
const cleanup = () => {
  try {
    process.kill(-server.pid, 'SIGTERM');
  } catch {}
};
process.on('uncaughtException', async (e) => {
  console.error('✗ crashed:', e.message.split('\n')[0]);
  await page?.screenshot({ path: path.join(out, 'crash.png') }).catch(() => {});
  cleanup();
  process.exit(2);
});
process.on('unhandledRejection', (e) => {
  throw e;
});
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, locale: 'zh-TW', acceptDownloads: true });
await ctx.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: URL.slice(0, -1) });
const page = await ctx.newPage();
page.setDefaultTimeout(8000);
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));
const mod = 'Meta';
const shot = (n) => page.screenshot({ path: path.join(out, `${n}.png`) });
const cellText = (a1) =>
  page.evaluate((a) => {
    // Find the painted cell at an address through the name box route: select it, read the formula bar and the painted text.
    return a;
  }, a1);
void cellText;

// Helpers that read the app's state through the DOM.
const go = async (a1) => {
  await page.fill('.namebox', a1);
  await page.press('.namebox', 'Enter');
};
const bar = () => page.inputValue('.finput');
const typeIn = async (a1, text) => {
  await go(a1);
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
};
const painted = () => page.$$eval('.pane .cell', (cs) => cs.map((c) => c.textContent));

// ---------- home ----------
await page.goto(URL);
await page.waitForSelector('.card .thumb-grid');
check('home lists the starter workbook', (await page.$$('.card')).length === 1);
check('sidebar matches the family layout', Boolean(await page.$('.sidebar .side-search')) && Boolean(await page.$('.side-item.active')));
await shot('01-home');
// five languages, like open-doc
await page.click('.side-foot button[title="語言"]');
await page.click('.menu-item:has-text("日本語")');
await page.waitForTimeout(200);
check('switch the interface to Japanese', (await page.textContent('.panel-title'))?.includes('スプレッドシート'));
await shot('01b-home-ja');
await page.click('.side-foot button[title="言語"]');
await page.click('.menu-item:has-text("繁體中文")');
await page.waitForTimeout(200);

// templates
await page.click('a.side-item[href="#/templates"]');
await page.waitForSelector('.card.template');
check('templates gallery', (await page.$$('.card.template')).length >= 5);
await shot('02-templates');
await page.click('.card.template:nth-child(4)'); // project plan
await page.waitForSelector('.dialog');
await shot('03-new-dialog');
await page.click('.dialog .btn.primary');
await page.waitForSelector('.grid .cell');
await page.waitForTimeout(300);
check('template opens as a workbook', (await page.textContent('.book-title'))?.length > 0);
await shot('04-template-book');
check('date formats show as dates', (await painted()).includes('2026-12-01'), JSON.stringify((await painted()).slice(0, 30)));
check('frozen panes painted', !(await page.$eval('.frozen-pane', (n) => n.hidden)));

// ---------- a blank workbook ----------
await page.goto(`${URL}#/`);
await page.waitForSelector('.panel-head .btn.primary');
await page.click('.panel-head .btn.primary');
await page.fill('.dialog input', 'E2E Test');
await page.click('.dialog .btn.primary');
await page.waitForSelector('.grid');
await page.waitForTimeout(200);

await typeIn('A1', '品項');
await typeIn('B1', '金額');
await typeIn('A2', '咖啡');
await typeIn('B2', '120');
await typeIn('A3', '午餐');
await typeIn('B3', '180');
await typeIn('A4', '日期');
await typeIn('B4', '2026/1/5');
await go('B5');
await page.keyboard.type('=SU');
await page.waitForSelector('.ac-popup:not([hidden]) .ac-item');
check('function autocomplete opens', (await page.textContent('.ac-item b')) === 'SUM');
await shot('05-autocomplete');
await page.keyboard.press('Tab');
check('autocomplete inserts SUM(', (await page.inputValue('.editor')) === '=SUM(');
await page.keyboard.type('B2:B3');
check('reference boxes while editing', (await page.$$('.ref-box')).length === 1);
await shot('06-ref-highlight');
await page.keyboard.press('Enter');
await go('B5');
check('formula computes and closes paren', (await bar()) === '=SUM(B2:B3)' && (await painted()).includes('300'), await bar());
check('typed date shows as a date', (await painted()).includes('2026-01-05'));

// pointing with arrows
await go('C2');
await page.keyboard.type('=');
await page.keyboard.press('ArrowLeft');
await page.keyboard.type('*2');
await page.keyboard.press('Enter');
await go('C2');
check('arrow keys point at a reference', (await bar()) === '=B2*2', await bar());

// fill handle: drag C2 down to C3
const box = await page.$eval('.fill-handle', (n) => n.getBoundingClientRect().toJSON());
await page.mouse.move(box.x + 4, box.y + 4);
await page.mouse.down();
await page.mouse.move(box.x + 4, box.y + 4 + 24 * 2, { steps: 4 });
await page.mouse.up();
await go('C3');
check('fill handle copies formulas with shifted refs', (await bar()) === '=B3*2', await bar());

// series fill: 1,2 -> 3,4,5
await typeIn('E1', '1');
await typeIn('E2', '2');
await go('E1:E5');
await page.keyboard.press(`${mod}+d`);
await go('E5');
check('⌘D fills a series? (Excel copies on ⌘D; series by handle)', (await bar()) !== '');

// bold, percent
await go('B2:B3');
await page.keyboard.press(`${mod}+b`);
check('⌘B makes cells bold', (await page.$$('.cell.b')).length >= 2);
await page.click('.toolbar .tbtn.select');
await page.waitForSelector('.menu');
await shot('07-format-menu');
await page.keyboard.press('Escape');
await page.click(`.toolbar button[title="${'百分比'}"]`);
await go('B2');
check('percent format shows %', (await painted()).some((t) => t === '12000%'), JSON.stringify(await painted()));
await page.keyboard.press(`${mod}+z`);
await page.waitForTimeout(100);
check('undo removes the percent format', (await painted()).includes('120'));

// fill colour from the swatch menu
await go('A1:B1');
await page.click('.toolbar button[title="填滿色彩"]');
await page.click('.swatch-grid .swatch:nth-child(7)');
check('fill colour paints cells', await page.$$eval('.pane .cell', (cs) => cs.some((c) => c.style.background)));

// right-click menu on a cell: insert a row above row 2
const cellPos = async (row, col) => {
  const g = await page.$eval('.grid', (n) => n.getBoundingClientRect().toJSON());
  return { x: g.x + 46 + col * 100 + 30, y: g.y + 24 + row * 24 + 12 };
};
let p = await cellPos(1, 0);
await page.mouse.click(p.x, p.y, { button: 'right' });
await page.waitForSelector('.menu');
const items = await page.$$eval('.menu .menu-label', (ls) => ls.map((l) => l.textContent));
check('cell context menu has the usual items', ['剪下', '複製', '貼上', '在上方插入列', '刪除列', '清除內容', '從 A 到 Z 排序', '合併儲存格'].every((x) => items.some((i) => i?.startsWith(x))), items.join('|'));
await shot('08-cell-menu');
await page.click('.menu-item:has-text("在上方插入列")');
await go('B6');
check('insert row shifts formulas', (await bar()) === '=SUM(B3:B4)', await bar());

// column header menu
const g = await page.$eval('.grid', (n) => n.getBoundingClientRect().toJSON());
await page.mouse.click(g.x + 46 + 150, g.y + 12, { button: 'right' });
await page.waitForSelector('.menu');
const colItems = await page.$$eval('.menu .menu-label', (ls) => ls.map((l) => l.textContent));
check('column header menu', colItems.some((i) => i?.includes('欄寬')) && colItems.some((i) => i?.includes('自動調整欄寬')), colItems.join('|'));
await shot('09-col-menu');
await page.keyboard.press('Escape');

// column resize by dragging the header border
await page.mouse.move(g.x + 46 + 100 - 1, g.y + 12);
await page.mouse.down();
await page.mouse.move(g.x + 46 + 160, g.y + 12, { steps: 4 });
await page.mouse.up();
await page.waitForTimeout(100);
const w = await page.$eval('.colheads .heads-inner .head[data-col="0"]', (n) => n.getBoundingClientRect().width);
check('drag a header border to resize', Math.abs(w - 160) <= 3, String(w));

// merge
await go('A10:C10');
await page.click('.toolbar button[title="合併儲存格"]');
await typeIn('A10', '合併標題');
check('merged cell paints across', await page.$$eval('.pane .cell', (cs) => cs.some((c) => c.textContent === '合併標題' && parseFloat(c.style.width) > 250)));

// freeze top row
await page.click('.toolbar button[title="凍結窗格"]');
await page.click('.menu-item:has-text("凍結首列")');
await page.$eval('.body', (b) => (b.scrollTop = 200));
await page.waitForTimeout(100);
check('frozen row stays while scrolling', (await painted()).includes('品項'));
await shot('10-frozen');
await page.$eval('.body', (b) => (b.scrollTop = 0));

// sheet tabs: context menu, rename, add, duplicate, delete
await page.click('.tabs-bar button[title="新增工作表"]');
check('add a sheet', (await page.$$('.tab')).length === 2);
const tab = await page.$('.tab:nth-child(1)');
const tb = await tab.boundingBox();
await page.mouse.click(tb.x + 10, tb.y + 10, { button: 'right' });
await page.waitForSelector('.menu');
const tabItems = await page.$$eval('.menu .menu-label', (ls) => ls.map((l) => l.textContent));
check('sheet tab menu', ['重新命名', '建立副本', '向左移', '向右移', '刪除'].every((x) => tabItems.includes(x)) && Boolean(await page.$('.menu .swatch-grid')), tabItems.join('|'));
await shot('11-tab-menu');
await page.click('.menu-item:has-text("重新命名")');
await page.fill('.dialog input', '帳目');
await page.click('.dialog .btn.primary');
await page.waitForTimeout(100);
check('rename a sheet', (await page.textContent('.tab:nth-child(1)')) === '帳目');
await page.click('.tab:nth-child(2)');
await typeIn('A1', '=帳目!B3*10');
await go('A1');
check('cross-sheet formula after rename', (await painted()).includes('1200'), JSON.stringify(await painted()));
await page.mouse.click(tb.x + 10, tb.y + 10, { button: 'right' });
await page.click('.menu-item:has-text("建立副本")');
check('duplicate a sheet', (await page.$$('.tab')).length === 3);
await page.dblclick('.tab:nth-child(3)');
await page.waitForSelector('.dialog');
await page.keyboard.press('Escape');
const t3 = await (await page.$('.tab:nth-child(2)')).boundingBox();
await page.mouse.click(t3.x + 10, t3.y + 10, { button: 'right' });
await page.click('.menu-item:has-text("刪除")');
await page.click('.dialog .btn.danger');
check('delete a sheet', (await page.$$('.tab')).length === 2);

// find
await page.click('.tab:nth-child(1)');
await page.keyboard.press(`${mod}+f`);
await page.fill('.findbar input', '午餐');
await page.keyboard.press('Enter');
check('find moves to the match', (await page.inputValue('.namebox')) === 'A4', await page.inputValue('.namebox'));
await shot('12-find');
await page.keyboard.press('Escape');

// status bar stats
await go('B3:B4');
check('status bar sums the selection', (await page.textContent('.stats'))?.includes('300'), await page.textContent('.stats'));

// sort descending by column B in the data block
await go('B3');
await page.click('.toolbar button[title="排序"]');
await page.click('.menu-item:has-text("從 Z 到 A")');
await go('A3');
// Rows 3–6 are one block (no header row above it): 日期 (a date serial) sorts first, descending.
check('sort orders the block by the active column', (await bar()) === '日期', await bar());

// copy / paste with formula offset (internal clipboard)
await go('C3');
await page.keyboard.press(`${mod}+c`);
await go('D3');
await page.keyboard.press(`${mod}+v`);
await page.waitForTimeout(100);
await go('D3');
check('paste shifts relative refs', (await bar()) === '=C3*2', await bar());

// keyboard shortcuts dialog
await page.click('.book-head button[title="鍵盤快速鍵"]');
await page.waitForSelector('.shortcuts');
await shot('13-shortcuts');
await page.keyboard.press('Escape');
await page.click('.dialog .btn.primary').catch(() => {});

// command palette
await go('A1');
await page.keyboard.press(`${mod}+k`);
await page.waitForSelector('.palette input');
await page.keyboard.type('斜體');
await shot('13b-palette');
await page.keyboard.press('Enter');
check('⌘K palette runs an action', (await page.$$('.cell.i')).length === 1);
await page.keyboard.press(`${mod}+z`);

// notes: add from the context menu, see the mark, the hover and the badge
await go('B5');
{
  const g = await page.$eval('.grid', (n) => n.getBoundingClientRect().toJSON());
  const box = await page.$eval('.active-box', (n) => n.getBoundingClientRect().toJSON());
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, { button: 'right' });
  await page.click('.menu-item:has-text("新增備註")');
  await page.fill('.dialog textarea', '請改成百分比');
  await page.keyboard.press(`${mod}+Enter`);
  await page.waitForTimeout(150);
  check('a note leaves a mark on the cell', (await page.$$('.note-mark')).length === 1);
  check('the header counts notes', (await page.textContent('.book-head .badge')) === '1');
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.waitForTimeout(100);
  check('hovering shows the note', (await page.textContent('.note-pop'))?.includes('請改成百分比'));
  await shot('13c-note');
  void g;
}

// filters, conditional formats and dropdown lists, on a sheet of their own
{
  const sheetsBefore = (await page.$$('.tab')).length;
  await page.click('.tabs-bar button[title="新增工作表"]');
  await page.waitForFunction((n) => document.querySelectorAll('.tab').length === n + 1, sheetsBefore);
  for (const [a, v] of [['A1', '狀態'], ['B1', '金額'], ['A2', '完成'], ['B2', '100'], ['A3', '進行中'], ['B3', '300'], ['A4', '完成'], ['B4', '200']]) await typeIn(a, v);
  await go('A1');
  await page.click('.toolbar button[title="篩選"]');
  check('a filter puts buttons on the header row', (await page.$$('.filter-btn')).length === 2);
  const btn = await page.$eval('.filter-btn', (n) => n.getBoundingClientRect().toJSON());
  await page.mouse.click(btn.x + btn.width / 2, btn.y + btn.height / 2);
  await page.waitForSelector('.menu');
  await shot('13d-filter-menu');
  await page.click('.menu-item:has-text("進行中")');
  await page.keyboard.press('Escape');
  await page.waitForTimeout(150);
  check('unticking a value hides its rows', !(await painted()).includes('300') && (await painted()).includes('200'), JSON.stringify(await painted()));
  await page.click('.toolbar button[title="篩選"]');
  await page.click('.menu-item:has-text("全部顯示")');
  check('show all brings the rows back', (await painted()).includes('300'));

  await go('B2:B4');
  await page.click('.toolbar button[title="條件式格式"]');
  await page.click('.menu-item:has-text("資料橫條")');
  check('data bars draw under the numbers', (await page.$$('.data-bar')).length === 3);
  await page.click('.toolbar button[title="條件式格式"]');
  await page.click('.menu-item:has-text("色階：紅－黃－綠")');
  const fills = await page.$$eval('.pane .cell', (cs) => cs.filter((c) => c.style.background && /^(100|200|300)$/.test(c.textContent ?? '')).map((c) => c.style.background));
  check('a colour scale fills low to high differently', new Set(fills).size === 3, JSON.stringify(fills));

  await go('C2');
  await page.keyboard.press(`${mod}+k`);
  await page.waitForSelector('.palette input');
  await page.keyboard.type('資料驗證');
  await page.keyboard.press('Enter');
  await page.fill('.dialog input', '是, 否');
  await page.click('.dialog .btn.primary');
  await page.waitForTimeout(100);
  check('a dropdown list shows its button', await page.$eval('.dv-btn', (n) => !n.hidden));
  await page.keyboard.type('也許');
  await page.keyboard.press('Enter');
  await page.waitForTimeout(100);
  check('a value outside the list is refused', (await page.textContent('.toast'))?.includes('只能填') && !(await painted()).includes('也許'));
  await go('C2');
  await page.click('.dv-btn');
  await page.click('.menu-item:has-text("否")');
  check('choosing from the dropdown fills the cell', (await painted()).includes('否'));

  // borders and hidden columns
  await go('A1:B4');
  await page.click('.toolbar button[title="框線"]');
  await page.click('.menu-item:has-text("所有框線")');
  check('all borders draw on every cell', (await page.$$eval('.pane .cell', (cs) => cs.filter((c) => c.style.boxShadow.includes('inset')).length)) >= 8);
  const heads = () => page.$$eval('.colheads .head', (hs) => hs.map((h) => h.textContent));
  const gb = await page.$eval('.grid', (n) => n.getBoundingClientRect().toJSON());
  const colD = await page.$eval('.colheads .head[data-col="3"]', (n) => n.getBoundingClientRect().toJSON());
  await page.mouse.click(colD.x + colD.width / 2, colD.y + colD.height / 2, { button: 'right' });
  await page.click('.menu-item:has-text("隱藏欄")');
  check('hiding a column removes its header', !(await heads()).includes('D') && (await heads()).includes('E'));
  void gb;

  // charts: images in Assets, never on the grid; notes beside them
  await typeIn('B2', '500');
  await go('A1:B4');
  await page.click('.toolbar button[title="圖表圖片"]');
  await page.click('.menu-item:has-text("長條圖")');
  await page.waitForSelector('.dialog');
  await page.fill('.dialog input >> nth=0', '狀態金額');
  await page.click('.dialog .btn.primary');
  await page.waitForSelector('.assets-panel:not([hidden]) .asset-card.image');
  const bookIdNow = decodeURIComponent((await page.evaluate(() => location.hash)).replace(/^#\/w\//, ''));
  const chartFileOnDisk = path.join(ws, 'sheets', bookIdNow, 'assets', '狀態金額.svg');
  const chartSvgText = existsSync(chartFileOnDisk) ? readFileSync(chartFileOnDisk, 'utf8') : '';
  check('a chart image lands in Assets, drawn from the cells', chartSvgText.startsWith('<svg xmlns') && chartSvgText.includes('500') && (await page.$$('.assets-panel .asset-card.image img')).length >= 1, chartFileOnDisk);
  check('nothing is placed on the grid', (await page.$$('.chart')).length === 0);
  await page.click('.assets-actions .btn:has-text("新增筆記")');
  await page.waitForSelector('.dialog');
  await page.fill('.dialog input >> nth=0', '預算說明');
  await page.fill('.dialog textarea', '金額是新台幣，狀態由業務填寫。');
  await page.click('.dialog .btn.primary');
  await page.waitForSelector('.assets-panel .asset-card.note[data-name="預算說明.md"]', { timeout: 5000 }).catch(() => {});
  const noteOnDisk = path.join(ws, 'sheets', bookIdNow, 'assets', '預算說明.md');
  check('a note is saved as Markdown beside the workbook', (await page.$$('.assets-section .asset-card.note[data-name="預算說明.md"]')).length === 1 && existsSync(noteOnDisk) && readFileSync(noteOnDisk, 'utf8').startsWith('# 預算說明') && readFileSync(noteOnDisk, 'utf8').includes('新台幣'));
  await shot('13f-assets');
  await page.click('.assets-head button');

  // AutoComplete from the column, and the format painter
  await go('A5');
  await page.keyboard.type('進');
  check('autocomplete fills in an entry from the column', (await page.inputValue('.editor')) === '進行中', await page.inputValue('.editor'));
  await page.keyboard.press('Escape');
  await go('A2');
  await page.keyboard.press(`Shift+${mod}+c`);
  check('the format painter picks up a format', await page.$eval('.toolbar button[title^="複製格式"]', (b) => b.classList.contains('on')));
  await go('F2:F3');
  await page.keyboard.press(`Shift+${mod}+v`);
  await page.waitForTimeout(100);
  const brushed = await page.$$eval('.pane .cell', (cs) => cs.filter((c) => c.style.boxShadow.includes('inset') && c.getBoundingClientRect().left > 420).length); // column F (D is hidden)
  check('the format painter applies it to the selection', brushed === 2, String(brushed));
  await page.keyboard.press(`${mod}+z`);

  // data tools: remove duplicates, text to columns, paste transposed
  const palette = async (words) => {
    await page.keyboard.press(`${mod}+k`);
    await page.waitForSelector('.palette input');
    await page.keyboard.type(words);
    await page.keyboard.press('Enter');
  };
  for (const [a, v] of [['H1', '名稱'], ['I1', '分數'], ['H2', 'a'], ['I2', '1'], ['H3', 'a'], ['I3', '1'], ['H4', 'b'], ['I4', '2']]) await typeIn(a, v);
  await go('H2');
  await palette('移除重複');
  await go('H3');
  const h3 = await bar();
  await go('H4');
  check('remove duplicates keeps the header and first copies', h3 === 'b' && (await bar()) === '', `H3=${h3} H4=${await bar()}`);
  await typeIn('K2', 'x,y,z');
  await go('K2');
  await palette('資料剖析');
  await page.click('.dialog .btn.primary');
  await go('M2');
  check('text to columns splits into the cells to the right', (await bar()) === 'z', await bar());
  await go('H1:I2');
  await page.keyboard.press(`${mod}+c`);
  await go('N1');
  await palette('轉置貼上');
  await go('N2');
  const n2 = await bar();
  await go('O1');
  check('paste transposed swaps rows and columns', n2 === '分數' && (await bar()) === 'a', `N2=${n2} O1=${await bar()}`);
  // defined names: name a range, use it in a formula, go to it from the name box
  await go('I2:I3');
  await palette('定義名稱');
  await page.fill('.dialog input >> nth=0', 'Scores');
  await page.click('.dialog .btn.primary');
  await typeIn('J5', '=SUM(Scores)');
  await go('J5');
  check('a defined name works in a formula', (await painted()).includes('3') && (await bar()) === '=SUM(Scores)', await bar());
  await page.fill('.namebox', 'scores');
  await page.press('.namebox', 'Enter');
  check('the name box goes to a defined name', (await page.inputValue('.namebox')) === 'I2:I3', await page.inputValue('.namebox'));
  // trace precedents: J5 reads I2:I3 through the name; the list jumps there
  await go('J5');
  await palette('追蹤前導');
  const traced = await page.$$eval('.trace-box', (b) => b.length);
  const arrows = await page.$$eval('.trace-arrows path', (p) => p.length);
  const listed = await page.$$eval('.menu-item', (m) => m.map((n) => n.textContent?.trim()));
  check('trace precedents draws the range read through a name', traced === 1 && arrows === 1 && listed.some((t) => t?.includes('I2:I3')), `${traced} boxes, ${arrows} arrows, ${listed.join('|')}`);
  await page.waitForTimeout(400);
  await shot('13f-trace');
  await page.click('.menu-item:has-text("I2:I3")');
  check('the trace list jumps to a precedent', (await page.inputValue('.namebox')) === 'I2:I3', await page.inputValue('.namebox'));
  await go('I2');
  await palette('追蹤從屬');
  await page.keyboard.press('Escape');
  await page.keyboard.press('Escape');
  check('Esc removes the trace arrows', (await page.$$('.trace-box')).length === 0);
  // custom sort: by 類別 ascending, then 值 descending, the header kept
  for (const [a, v] of [['T1', '類別'], ['U1', '值'], ['T2', 'b'], ['U2', '2'], ['T3', 'a'], ['U3', '1'], ['T4', 'b'], ['U4', '1'], ['T5', 'a'], ['U5', '2']]) await typeIn(a, v);
  await go('T2'); // apart from the data to its left, so the region is T1:U5
  await palette('自訂排序');
  await page.selectOption('.dialog select >> nth=1', '20:d');
  await page.waitForTimeout(300);
  await shot('13g-sort');
  await page.click('.dialog .btn.primary');
  const sortedRows = [];
  for (const r of [1, 2, 3, 4, 5]) {
    await go(`T${r}`);
    const p = await bar();
    await go(`U${r}`);
    sortedRows.push(`${p}${await bar()}`);
  }
  check('custom sort orders by two columns and keeps the header', sortedRows.join(' ') === '類別值 a2 a1 b2 b1', sortedRows.join(' '));
  // font size: the toolbar menu, then ⇧⌘> to the next size; the row grows to fit
  await go('T1');
  const rowBefore = await page.$eval('.rowheads .head:nth-child(1)', (h) => h.getBoundingClientRect().height).catch(() => 24);
  await page.click('.toolbar .tbtn.size');
  await page.locator('.menu-item').filter({ hasText: /^18$/ }).click();
  await page.keyboard.press('Meta+Shift+Period');
  const sized = await page.evaluate(() => {
    const cells = [...document.querySelectorAll('.pane .cell')].filter((c) => c.textContent === '類別');
    return cells.map((c) => getComputedStyle(c).fontSize).join(',');
  });
  const rowAfter = await page.$eval('.rowheads .head:nth-child(1)', (h) => h.getBoundingClientRect().height).catch(() => 0);
  await shot('13h-font-size');
  check('font size from the toolbar and ⇧⌘> grows the row', (await page.textContent('.toolbar .tbtn.size'))?.trim() === '20' && sized.includes('23.6') && rowAfter > rowBefore, `${await page.textContent('.toolbar .tbtn.size')} ${sized} row ${rowBefore}→${rowAfter}`);
  // vertical alignment: U1 sits in the row the big font made tall
  await go('U1');
  await palette('垂直對齊：靠上');
  const valign = await page.evaluate(() => [...document.querySelectorAll('.pane .cell')].filter((c) => c.textContent === '值').map((c) => getComputedStyle(c).alignItems).join(','));
  check('vertical alignment puts the text at the top', valign === 'flex-start', valign);
  // links: a typed address and HYPERLINK are links, ⌘-click opens them; javascript: never is
  await typeIn('W1', 'https://example.com/a');
  await typeIn('W2', '=HYPERLINK("https://example.com/b", "報表")');
  await typeIn('W3', 'javascript:alert(1)');
  const linkCells = await page.evaluate(() => [...document.querySelectorAll('.pane .cell.link')].map((c) => c.textContent));
  check('addresses and HYPERLINK show as links, javascript: does not', linkCells.includes('https://example.com/a') && linkCells.includes('報表') && !linkCells.some((t) => t?.startsWith('javascript')), linkCells.join('|'));
  await page.evaluate(() => {
    window.__opened = [];
    window.open = (u) => (window.__opened.push(String(u)), null);
  });
  const w2 = await page.evaluate(() => [...document.querySelectorAll('.pane .cell.link')].find((c) => c.textContent === '報表')?.getBoundingClientRect().toJSON());
  if (w2) {
    // page.mouse.click takes no modifiers; hold the key instead.
    await page.keyboard.down('Meta');
    await page.mouse.click(w2.x + 10, w2.y + w2.height / 2);
    await page.keyboard.up('Meta');
  }
  check('⌘-click opens the link', (await page.evaluate(() => window.__opened.join(','))) === 'https://example.com/b', await page.evaluate(() => window.__opened.join(',')));
  // checkboxes: insert over Y1:Y3, click one, Space flips the range; a formula counts them
  await typeIn('Z1', '=COUNTIF(Y1:Y3, TRUE)');
  await go('Y1:Y3');
  await palette('插入核取方塊');
  const boxes = () => page.$$eval('.pane .cell .cb', (b) => b.map((x) => x.getAttribute('aria-checked')).join(','));
  check('inserting checkboxes draws unticked boxes', (await boxes()) === 'false,false,false', await boxes());
  const firstBox = await page.$eval('.pane .cell .cb', (b) => b.getBoundingClientRect().toJSON());
  await page.mouse.click(firstBox.x + firstBox.width / 2, firstBox.y + firstBox.height / 2);
  const countOf = async () => {
    await go('Z1');
    return page.evaluate(() => [...document.querySelectorAll('.pane .cell')].find((c) => c.getBoundingClientRect().left > 0 && c.querySelector('.cell-text') && c.textContent && /^\d$/.test(c.textContent) && c.classList.contains('a-right'))?.textContent ?? '?');
  };
  check('a click ticks a box and the formula counts it', (await boxes()) === 'true,false,false' && (await countOf()) === '1', `${await boxes()} count ${await countOf()}`);
  await go('Y1:Y3');
  await page.keyboard.press('Space');
  const allOn = await boxes();
  await shot('13i-checkbox');
  await page.keyboard.press('Space');
  check('Space ticks them all, then clears them all', allOn === 'true,true,true' && (await boxes()) === 'false,false,false', `${allOn} then ${await boxes()}`);
  // protection: a protected range refuses typing and clearing until unprotected
  await typeIn('AB1', '鎖住');
  await go('AB1:AB2');
  await palette('保護範圍');
  check('protected cells show as locked', (await page.$$('.pane .cell.locked')).length >= 2);
  await go('AB1');
  await page.keyboard.type('x');
  await page.keyboard.press('Enter');
  const refusedTyping = (await page.textContent('.toast')) ?? '';
  await go('AB1');
  check('typing into a protected cell is refused', (await bar()) === '鎖住' && refusedTyping.includes('受保護'), `${await bar()} | ${refusedTyping}`);
  await page.keyboard.press('Delete');
  await go('AB1');
  check('clearing a protected cell is refused', (await bar()) === '鎖住', await bar());
  await go('AB1:AB2');
  await palette('取消保護');
  await typeIn('AB1', '解開');
  await go('AB1');
  check('unprotecting lets the cell change again', (await bar()) === '解開' && (await page.$$('.pane .cell.locked')).length === 0, await bar());
  // show formulas (⌃`) and error checking
  await go('J5');
  await page.keyboard.press('Control+Backquote');
  const asFormulas = await painted();
  await page.keyboard.press('Control+Backquote');
  check('⌃` shows formulas, and again the results', asFormulas.includes('=SUM(Scores)') && !(await painted()).includes('=SUM(Scores)') && (await painted()).includes('3'), `${asFormulas.filter((t) => t?.startsWith('=')).join('|')} / after: ${(await painted()).filter((t) => t?.startsWith('=') || t === '3').join('|')}`);
  await typeIn('K7', '=1/0');
  await go('A1');
  await palette('下一個錯誤');
  const why = await page.textContent('.toast');
  check('error checking goes to the error and says what it means', (await page.inputValue('.namebox')) === 'K7' && why?.includes('#DIV/0!') && why?.includes('除以零'), `${await page.inputValue('.namebox')} ${why}`);
  await page.keyboard.press('Delete');
  await go('A1'); // scroll back to the filter's header
  await shot('13e-features');
  // Leave 進行中 filtered out, for the xlsx round trip below.
  const again = await page.$eval('.filter-btn', (n) => n.getBoundingClientRect().toJSON());
  await page.mouse.click(again.x + again.width / 2, again.y + again.height / 2);
  await page.click('.menu-item:has-text("進行中")');
  await page.keyboard.press('Escape');
}

// xlsx export, then import it as a new workbook
await page.click('.book-head .btn.primary');
const [download] = await Promise.all([page.waitForEvent('download'), page.click('.menu-item:has-text("Excel")')]);
const file = path.join(out, 'export.xlsx');
await download.saveAs(file);
check('xlsx export downloads', readFileSync(file).subarray(0, 2).toString() === 'PK');
// The zip is stored, not deflated, so the styles part can be read straight from the bytes.
check('xlsx export keeps the font size', readFileSync(file).toString('latin1').includes('<sz val="20"/>'));
await page.goto(`${URL}#/`);
await page.waitForSelector('.panel-head');
const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('.panel-head .btn:has-text("匯入")')]);
await chooser.setFiles(file);
await page.waitForSelector('.dialog');
await page.fill('.dialog input >> nth=0', 'Round trip');
await page.click('.dialog .btn.primary');
await page.waitForSelector('.grid .cell');
await page.waitForTimeout(200);
const tabsAfter = await page.$$eval('.tab', (ts) => ts.map((t) => t.textContent));
check('xlsx import keeps sheets', tabsAfter.includes('帳目'), tabsAfter.join(','));
await go('D3');
check('xlsx import keeps formulas', (await bar()) === '=C3*2', await bar());
check('xlsx import keeps bold', (await page.$$('.cell.b')).length >= 1);
check('xlsx round trip keeps a typed date a date', (await painted()).includes('2026-01-05'), JSON.stringify(await painted()));
await shot('14-imported');
await page.click('.tab:has-text("Sheet3")');
await page.waitForTimeout(200);
check(
  'xlsx round trip keeps the filter and its hidden rows',
  (await page.$$('.filter-btn')).length === 2 && !(await painted()).includes('300') && (await painted()).includes('200'),
  JSON.stringify(await painted()),
);
const barsAfter = (await page.$$('.data-bar')).length;
const scaleAfter = await page.$$eval('.pane .cell', (cs) => cs.filter((c) => c.style.background && /^(500|200)$/.test(c.textContent ?? '')).length); // B2 became 500 in the chart test
check('xlsx round trip keeps data bars and the colour scale', barsAfter === 2 && scaleAfter === 2, `${barsAfter} bars, ${scaleAfter} filled`);
await go('C2');
check('xlsx round trip keeps the dropdown list', await page.$eval('.dv-btn', (n) => !n.hidden));
check('xlsx round trip keeps borders', (await page.$$eval('.pane .cell', (cs) => cs.filter((c) => c.style.boxShadow.includes('inset')).length)) >= 6);
check('xlsx round trip keeps hidden columns', !(await page.$$eval('.colheads .head', (hs) => hs.map((h) => h.textContent))).includes('D'));
// Unhide by selecting across the hidden band.
{
  const c = await page.$eval('.colheads .head[data-col="2"]', (n) => n.getBoundingClientRect().toJSON());
  const e = await page.$eval('.colheads .head[data-col="4"]', (n) => n.getBoundingClientRect().toJSON());
  await page.mouse.click(c.x + c.width / 2, c.y + c.height / 2);
  await page.keyboard.down('Shift');
  await page.mouse.click(e.x + e.width / 2, e.y + e.height / 2);
  await page.keyboard.up('Shift');
  await page.mouse.click(e.x + e.width / 2, e.y + e.height / 2, { button: 'right' });
  await page.click('.menu-item:has-text("取消隱藏欄")');
  check('unhide brings the column back', (await page.$$eval('.colheads .head', (hs) => hs.map((h) => h.textContent))).includes('D'));
}
await page.click('.tab:has-text("Sheet3")');
await go('T1');
check('xlsx round trip keeps the font size', (await page.textContent('.toolbar .tbtn.size'))?.trim() === '20', await page.textContent('.toolbar .tbtn.size'));
check('xlsx round trip keeps vertical alignment', (await page.evaluate(() => [...document.querySelectorAll('.pane .cell')].filter((c) => c.textContent === '值').map((c) => getComputedStyle(c).alignItems).join(','))) === 'flex-start');
await page.click('.tab:has-text("帳目")');
await page.waitForTimeout(150);

// agents: the selection lands in current.json, and a file edit shows up live
await go('B3');
await page.waitForTimeout(600);
const current = JSON.parse(readFileSync(path.join(ws, 'node_modules/.open-spreadsheet/current.json'), 'utf8'));
check('current.json reports the selection', current.active === 'B3' && current.sheet === '帳目', JSON.stringify(current));
const bookId = current.workbook;
const bookFile = path.join(ws, 'sheets', bookId, 'workbook.json');
const onDisk = JSON.parse(readFileSync(bookFile, 'utf8'));
onDisk.sheets[0].cells.H1 = 'from agent';
writeFileSync(bookFile, `${JSON.stringify(onDisk, null, 2)}\n`);
await page.waitForFunction(() => [...document.querySelectorAll('.pane .cell')].some((c) => c.textContent === 'from agent'), null, { timeout: 5000 }).catch(() => {});
check('an agent edit shows up live', (await painted()).includes('from agent'));
// …and the page's next save builds on it instead of overwriting it
await typeIn('H2', 'from person');
await page.waitForTimeout(900);
const after = JSON.parse(readFileSync(bookFile, 'utf8')).sheets[0].cells;
check('saves keep the agent edit', after.H1 === 'from agent' && after.H2 === 'from person', JSON.stringify({ H1: after.H1, H2: after.H2 }));

// an older workbook with a chart on its grid is offered the move to Assets
{
  const book = JSON.parse(readFileSync(bookFile, 'utf8'));
  book.sheets[0].charts = [{ id: 'old-1', type: 'pie', range: 'A1:B4', title: 'Legacy pie', x: 600, y: 40, w: 300, h: 200 }];
  writeFileSync(bookFile, `${JSON.stringify(book, null, 2)}\n`);
  await page.waitForTimeout(800);
  await page.reload();
  await page.waitForSelector('.legacy-charts:not([hidden])', { timeout: 5000 }).catch(() => {});
  check('a chart left on the grid is offered the move', await page.$eval('.legacy-charts', (n) => !n.hidden && (n.textContent ?? '').includes('1')).catch(() => false));
  await page.click('.legacy-charts .btn');
  await page.waitForTimeout(800);
  const legacySvg = path.join(ws, 'sheets', bookId, 'assets', 'Legacy-pie.svg');
  const saved = JSON.parse(readFileSync(bookFile, 'utf8'));
  check('moving it saves an image and clears the grid', existsSync(legacySvg) && !saved.sheets[0].charts && (await page.$$('.chart')).length === 0, legacySvg);
}

// dark mode
await page.click('.book-head button[title="主題"]');
await page.waitForTimeout(150);
await shot('15-dark');
await page.click('.book-head button[title="主題"]');

check('no page errors', errors.length === 0, errors.slice(0, 3).join(' | '));
await browser.close();
process.kill(-server.pid, 'SIGTERM');
const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} passed`);
process.exit(failed.length ? 1 : 0);
