import { mountAgentPanel } from './agent-panel.ts';
import { openBook } from './book.ts';
import { renderHome } from './home.ts';
import { applyTheme, el } from './ui.ts';

applyTheme();
// The page on the left; the agent panel, when shown, on the right.
const layout = el('div', 'app-layout');
const app = el('div', 'app-main');
layout.append(app);
document.getElementById('app')!.append(layout);
mountAgentPanel(layout);

async function route(): Promise<void> {
  const book = /^#\/w\/([^/]+)$/.exec(location.hash);
  if (book) return openBook(app, decodeURIComponent(book[1]!));
  return renderHome(app, location.hash === '#/templates' ? 'templates' : 'books');
}

window.addEventListener('hashchange', () => void route());
void route();
