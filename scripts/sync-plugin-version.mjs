#!/usr/bin/env node
/**
 * Gives the agent plugin core's version. Installed plugins update only when
 * their version changes, so it follows the package's (agent.test.ts checks
 * it). Runs after `changeset version` in `pnpm version-packages`.
 */
import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const { version } = JSON.parse(readFileSync(path.join(root, 'packages/core/package.json'), 'utf8'));
for (const file of ['plugin.json', '.claude-plugin/plugin.json']) {
  const target = path.join(root, 'plugins/open-spreadsheet', file);
  const text = readFileSync(target, 'utf8');
  writeFileSync(target, text.replace(/("version":\s*")[^"]*(")/, `$1${version}$2`));
}
process.stdout.write(`plugin version → ${version}\n`);
