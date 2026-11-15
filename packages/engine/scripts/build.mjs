#!/usr/bin/env node
// Writes dist/: the engine's sources with their types stripped and `.ts`
// imports pointed at `.js`. Node will not strip types inside node_modules, so
// the published package has to ship JavaScript; in this repo everything still
// runs straight from src/.
import { mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { stripTypeScriptTypes } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

process.removeAllListeners('warning'); // stripTypeScriptTypes is marked experimental
const pkg = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = process.argv[2] ? path.resolve(process.argv[2]) : path.join(pkg, 'dist');
rmSync(out, { recursive: true, force: true });
mkdirSync(out, { recursive: true });
const SPEC = /((?:^|\n)\s*(?:import|export)\b[^;]*?['"])(\.{1,2}\/[^'"]+)\.ts(['"])/g;
let n = 0;
for (const f of readdirSync(path.join(pkg, 'src'))) {
  if (!f.endsWith('.ts') || f.endsWith('.test.ts')) continue;
  const js = stripTypeScriptTypes(readFileSync(path.join(pkg, 'src', f), 'utf8'), { mode: 'strip' }).replace(SPEC, '$1$2.js$3');
  writeFileSync(path.join(out, f.replace(/\.ts$/, '.js')), js);
  n++;
}
console.log(`engine: ${n} files -> ${path.relative(process.cwd(), out) || '.'}`);
