#!/usr/bin/env node
// Writes dist/charts.js: the chart renderer the editor uses, with its types
// stripped, so the MCP server and CLI can draw a chart image from an installed
// package (Node will not strip types inside node_modules).
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { stripTypeScriptTypes } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

process.removeAllListeners('warning'); // stripTypeScriptTypes is marked experimental
const pkg = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const out = process.argv[2] ? path.resolve(process.argv[2]) : path.join(pkg, 'dist');
mkdirSync(out, { recursive: true });
const js = stripTypeScriptTypes(readFileSync(path.join(pkg, 'app', 'charts.ts'), 'utf8'), { mode: 'strip' });
if (/from '\.\.?\//.test(js)) throw new Error('charts.ts must import only @open-spreadsheet/engine at runtime');
writeFileSync(path.join(out, 'charts.js'), js);
console.log(`charts -> ${path.relative(process.cwd(), path.join(out, 'charts.js')) || 'charts.js'}`);
