import { openBook } from './book.ts';
import { renderHome } from './home.ts';
import { applyTheme } from './ui.ts';

applyTheme();
const app = document.getElementById('app')!;

async function route(): Promise<void> {
  const book = /^#\/w\/([^/]+)$/.exec(location.hash);
  if (book) return openBook(app, decodeURIComponent(book[1]!));
  return renderHome(app, location.hash === '#/templates' ? 'templates' : 'books');
}

window.addEventListener('hashchange', () => void route());
void route();
