import { readdir, readFile } from 'node:fs/promises';

const MAX_WORDS = 45;
const DIR = '.changeset';
const SKIP = new Set(['README.md', 'config.json']);

function body(text: string): string {
  const lines = text.split(/\r?\n/);
  if (lines[0] !== '---') return text;
  const end = lines.indexOf('---', 1);
  return end === -1 ? text : lines.slice(end + 1).join('\n');
}

const files = (await readdir(DIR)).filter((file) => file.endsWith('.md') && !SKIP.has(file));
let failed = false;
for (const file of files) {
  const text = await readFile(`${DIR}/${file}`, 'utf8');
  const words = body(text).trim().split(/\s+/).filter(Boolean).length;
  if (words > MAX_WORDS) {
    console.error(
      `${file}: ${words}-word body, over the ${MAX_WORDS}-word tl;dr limit (AGENTS.md, "Changesets").`
    );
    failed = true;
  }
}
if (failed) process.exit(1);
