import { readFileSync, readdirSync } from 'node:fs';
import { resolve, relative, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

const root = resolve(fileURLToPath(new URL('../..', import.meta.url)));
const allowed = /^(?:[a-z0-9-]+\.)?example\.(?:invalid|com|org|net)$/i;
function violations(text) {
  return [...text.matchAll(/[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@([a-z0-9-]+(?:\.[a-z0-9-]+)+)/gi)]
    .filter(match => !allowed.test(match[1])).length;
}
function files(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap(entry => {
    if (['bin', 'obj', 'node_modules'].includes(entry.name)) return [];
    const path = resolve(directory, entry.name);
    return entry.isDirectory() ? files(path) : /\.(?:cs|mjs|js|html|json|ps1|txt)$/i.test(extname(path)) ? [path] : [];
  });
}
// Negative control is deliberately in memory: never print addresses or fixture content.
if (process.argv.includes('--baseline')) {
  const original = execFileSync('git', ['show', 'main:tests/harness/addin-compose-segmentation.html'], { cwd: root, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
  const count = violations(original);
  console.log(`BASELINE: ${count} non-synthetic addresses rejected (values withheld).`);
  process.exit(count > 0 ? 1 : 0);
}
const targets = [...files(resolve(root, 'tests')), ...files(resolve(root, 'tools/Maily.Generation.TestRunner'))];
let total = 0;
for (const path of targets) {
  const count = violations(readFileSync(path, 'utf8'));
  if (count) console.error(`FAIL: ${relative(root, path)} has ${count} non-synthetic addresses (values withheld).`);
  total += count;
}
console.log(`Fixture safety: ${targets.length} files checked; ${total} non-synthetic addresses.`);
process.exit(total ? 1 : 0);
