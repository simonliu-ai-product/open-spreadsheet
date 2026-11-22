// End-to-end check of the agent panel in Chromium, with a scripted stand-in
// for @openai/codex-sdk installed in a fresh workspace: no Codex login, no API
// key, no network. Drives the toggle and ⌘J, a turn that edits a workbook (the
// grid reloads), the check after a turn, questions, history, a narrow window
// and a host that has its own agent, and saves screenshots.
//
//   node scripts/agent-e2e.mjs <out dir>
import { spawn } from 'node:child_process';
import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const out = path.resolve(process.argv[2] ?? 'agent-e2e-out');
mkdirSync(out, { recursive: true });
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const bin = path.join(root, 'packages/core/bin.js');
const ws = path.join(mkdtempSync(path.join(tmpdir(), 'oss-agent-e2e-')), 'book');
await new Promise((res, rej) => spawn(process.execPath, [bin, 'init', ws], { stdio: 'ignore' }).on('exit', (c) => (c ? rej(new Error('init')) : res())));

// The stand-in SDK: what the panel resolves from the workspace in place of the real one.
const sdk = path.join(ws, 'node_modules', '@openai', 'codex-sdk');
mkdirSync(sdk, { recursive: true });
writeFileSync(path.join(sdk, 'package.json'), JSON.stringify({ name: '@openai/codex-sdk', version: '0.0.0-e2e', type: 'module', exports: { '.': { import: './index.js' } } }));
writeFileSync(
  path.join(sdk, 'index.js'),
  `import { readFileSync, renameSync, writeFileSync } from 'node:fs';
import path from 'node:path';
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
function setCells(cwd, cells) {
  const file = path.join(cwd, 'sheets', 'getting-started', 'workbook.json');
  const data = JSON.parse(readFileSync(file, 'utf8'));
  Object.assign(data.sheets[0].cells, cells);
  writeFileSync(file + '.tmp', JSON.stringify(data, null, 2) + '\\n');
  renameSync(file + '.tmp', file);
}
export class Codex {
  startThread(o) { return this.thread('th-e2e', o); }
  resumeThread(id, o) { return this.thread(id, o); }
  thread(id, o) {
    return { id, async runStreamed(input) {
      globalThis.__agentInputs = [...(globalThis.__agentInputs ?? []), input];
      const said = input.split('<open-spreadsheet-context>')[0];
      return { events: (async function* () {
        yield { type: 'thread.started', thread_id: id };
        if (said.includes('ASK')) {
          yield { type: 'item.completed', item: { id: 'a-ask', type: 'agent_message', text: 'One question first.\\n<open-spreadsheet-questions>[{"header":"Total","question":"Which total?","options":[{"label":"Sum"},{"label":"Average"}]}]</open-spreadsheet-questions>' } };
        } else if (said.includes('My answers')) {
          yield { type: 'item.completed', item: { id: 'a-ans', type: 'agent_message', text: 'Using **Sum**, as you picked.' } };
        } else {
          const broken = said.includes('BREAK');
          const call = { id: 'm-' + Date.now(), type: 'mcp_tool_call', server: 'open-spreadsheet', tool: 'write_cells', arguments: { workbook: 'getting-started', sheet: 'Sheet1', cells: { A5: 'Double', B5: '' } }, status: 'in_progress' };
          yield { type: 'item.started', item: call };
          await wait(300);
          setCells(o.workingDirectory, broken ? { A5: 'Double', B5: '=NOPE(B4)' } : { A5: 'Double', B5: '=B4*2' });
          yield { type: 'item.completed', item: { ...call, status: 'completed' } };
          yield { type: 'item.completed', item: { id: 'a-' + Date.now(), type: 'agent_message', text: broken ? 'Wrote a formula.' : 'Added **Double** in B5: \`=B4*2\`.' } };
        }
        yield { type: 'turn.completed', usage: {} };
      })() };
    } };
  }
}
`,
);

const PORT = 5397;
const server = spawn(process.execPath, [bin, 'dev', '--port', String(PORT)], { cwd: ws, stdio: ['ignore', 'pipe', 'inherit'], detached: true });
await new Promise((res) => server.stdout.on('data', (d) => String(d).includes('Local') && res()));
const URL = `http://localhost:${PORT}/`;

const results = [];
const check = (name, ok, detail = '') => {
  results.push({ name, ok: Boolean(ok) });
  console.log(`${ok ? '✓' : '✗'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const browser = await chromium.launch();
const cleanup = () => {
  try {
    process.kill(-server.pid, 'SIGTERM');
  } catch {}
};
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, locale: 'en-US' });
await ctx.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: URL.slice(0, -1) });
const page = await ctx.newPage();
page.setDefaultTimeout(8000);
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));
const shot = (n, p = page) => p.screenshot({ path: path.join(out, `${n}.png`) });
const painted = () => page.$$eval('.pane .cell', (cs) => cs.map((c) => c.textContent));
const lastItems = () => page.$$eval('.agent-items > li', (ls) => ls.map((l) => l.dataset.kind));

try {
  await page.goto(`${URL}#/w/getting-started`);
  await page.waitForSelector('.grid .cell');
  check('the header has the agent toggle', await page.isVisible('.book-head .agent-toggle'));
  check('the panel starts closed', !(await page.isVisible('.agent-panel')));
  await page.keyboard.press('Meta+j');
  await page.waitForSelector('.agent-panel:not([hidden]) .agent-quick-btn');
  check('⌘J opens the panel on the empty state', true);
  check('the composer has the focus', await page.evaluate(() => document.activeElement?.classList.contains('agent-input')));
  await page.waitForFunction(() => document.querySelector('.agent-chip')?.textContent?.includes('Sheet1!A1')).catch(() => {});
  check('the context chip names the workbook and selection', (await page.textContent('.agent-chip'))?.includes('Getting started · Sheet1!A1'));
  await shot('01-empty');

  // A turn that writes a cell: the transcript streams it, the grid reloads, the check runs.
  await page.fill('.agent-input', 'Add a doubled total');
  await page.keyboard.press('Enter');
  await page.waitForSelector('.agent-items li[data-kind="check"] .agent-check-line');
  const kinds = await lastItems();
  check('the turn shows message, tool, reply, result and check', JSON.stringify(kinds) === JSON.stringify(['user', 'tools', 'assistant', 'result', 'check']), JSON.stringify(kinds));
  check('the tool row says what it touched', (await page.textContent('.agent-tool'))?.includes('write_cells') && (await page.textContent('.agent-tool-detail'))?.includes('getting-started · Sheet1 · A5, B5'));
  check('the reply renders Markdown', (await page.innerHTML('.agent-md')).includes('<strong>Double</strong>'));
  check('a clean check says so', (await page.textContent('.agent-check-line'))?.includes('no problems'));
  await page.waitForFunction(() => [...document.querySelectorAll('.pane .cell')].some((c) => c.textContent === '600'));
  check('the grid reloaded the agent’s change', (await painted()).includes('Double'));
  const inputs = await page.evaluate(() => fetch('/__agent/sessions').then((r) => r.json()));
  check('the session is tied to the workbook', inputs.sessions[0]?.workbook === 'getting-started');
  await shot('02-turn');

  // A turn that leaves a broken formula: the check lists it and offers a fix.
  await page.fill('.agent-input', 'BREAK something');
  await page.keyboard.press('Enter');
  await page.waitForSelector('.agent-check .agent-findings li.error');
  check('the check after the turn lists the broken formula', (await page.textContent('.agent-check'))?.includes('Sheet1!B5'));
  check('the newest check offers a fix', await page.isVisible('.agent-check .btn.primary'));
  await shot('03-check-found');

  // Questions as a form; the answer starts the next turn.
  await page.fill('.agent-input', 'ASK me');
  await page.keyboard.press('Enter');
  await page.waitForSelector('.agent-questions');
  check('Continue waits for an answer', await page.isDisabled('.agent-questions .btn.primary'));
  await page.click('.agent-option:has-text("Sum")');
  await shot('04-question');
  await page.click('.agent-questions .btn.primary');
  await page.waitForSelector('.agent-md:has-text("as you picked")');
  check('answering continues the conversation', (await page.textContent('.agent-answered'))?.includes('Sum'));

  // Model menu, history and a new chat.
  await page.click('.agent-model');
  check('the model menu offers Manage agents', await page.isVisible('.menu-item:has-text("Manage agents")'));
  await page.click('.menu-item:has-text("Manage agents")');
  await page.waitForSelector('.agent-card');
  check('the connect view shows the SDK as connected', (await page.textContent('.agent-card'))?.includes('Connected'));
  await shot('05-connect');
  await page.click('.agent-text-btn');
  await page.click('.agent-head button[aria-label="History"]');
  await page.waitForSelector('.agent-history-item');
  check('history lists the conversation under this workbook', (await page.textContent('.agent-history'))?.includes('This workbook'));
  await shot('06-history');
  await page.click('.agent-head button[aria-label="History"]');
  await page.click('.agent-head button[aria-label="New chat"]');
  await page.waitForSelector('.agent-quick-btn');
  check('a new chat starts empty', true);

  // The panel survives going home, and ⌘J closes it.
  await page.goto(`${URL}#/`);
  await page.waitForSelector('.card');
  check('home has the toggle and the panel stays open', (await page.isVisible('.side-foot .agent-toggle')) && (await page.isVisible('.agent-panel')));
  check('home offers creating a workbook', (await page.textContent('.agent-quick'))?.includes('Create a workbook'));
  await page.keyboard.press('Meta+j');
  check('⌘J closes it', !(await page.isVisible('.agent-panel')));

  // Narrow: the header still fits, and the panel covers the page.
  await page.setViewportSize({ width: 480, height: 800 });
  await page.goto(`${URL}#/w/getting-started`);
  await page.waitForSelector('.grid .cell');
  const fits = await page.$eval('.book-head', (h) => h.scrollWidth <= h.clientWidth + 1);
  check('a narrow header still fits with the toggle', fits && (await page.isVisible('.book-head .agent-toggle')));
  await page.click('.book-head .agent-toggle');
  await page.waitForSelector('.agent-panel:not([hidden])');
  const box = await page.$eval('.agent-panel', (p) => p.getBoundingClientRect().width);
  check('a narrow window gets the panel over the page', box >= 400 && box <= 480, String(box));
  await shot('07-narrow');
  await page.click('.agent-head button[aria-label="Close the agent panel"]');

  // In Traditional Chinese.
  await page.setViewportSize({ width: 1280, height: 800 });
  await page.evaluate(() => localStorage.setItem('open-spreadsheet:lang', 'zh'));
  await page.reload();
  await page.waitForSelector('.grid .cell');
  await page.keyboard.press('Meta+j');
  await page.waitForSelector('.agent-panel:not([hidden])');
  check('the panel speaks the interface language', (await page.getAttribute('.agent-head button[aria-label="關閉 Agent 面板"]', 'title')) === '關閉 Agent 面板');
  await shot('08-zh');
  await page.evaluate(() => localStorage.setItem('open-spreadsheet:lang', 'en'));

  // Dark, with the earlier conversation picked from History.
  await page.evaluate(() => localStorage.setItem('open-spreadsheet:theme', 'dark'));
  await page.reload();
  await page.waitForSelector('.grid .cell');
  await page.click('.agent-head button[aria-label="History"]');
  await page.click('.agent-history-item >> nth=-1');
  await page.waitForSelector('.agent-items li[data-kind="check"]');
  check('a conversation reopens from History', (await lastItems()).length > 5);
  await shot('09-dark');
  await page.evaluate(() => localStorage.removeItem('open-spreadsheet:theme'));

  // A host with its own agent: no panel, no toggle, no shortcut.
  const hosted = await ctx.newPage();
  await hosted.addInitScript(() => {
    window.openxHost = { protocol: 1, name: 'test host', capabilities: ['theme', 'agent'] };
    localStorage.setItem('open-spreadsheet:agent-open', '1');
  });
  await hosted.goto(`${URL}#/w/getting-started`);
  await hosted.waitForSelector('.grid .cell');
  await hosted.keyboard.press('Meta+j');
  check('a host with its own agent hides the panel and its toggle', !(await hosted.$('.agent-panel')) && !(await hosted.$('.agent-toggle')));
  await hosted.close();

  check('no page errors', errors.length === 0, errors.slice(0, 3).join(' | '));
} catch (e) {
  check('crashed', false, String(e.message).split('\n')[0]);
  await shot('crash').catch(() => {});
} finally {
  await browser.close();
  cleanup();
}
const failed = results.filter((r) => !r.ok).length;
console.log(`\n${results.length - failed}/${results.length} passed · screenshots in ${out}`);
process.exit(failed ? 1 : 0);
