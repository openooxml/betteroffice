import { createHash } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fixture } from '../test/fixture';

const root = resolve(process.argv[2] ?? '/tmp/betteroffice-agent-eval');
await mkdir(root, { recursive: true });
const hashes: Record<string, string> = {};
for (const name of ['revenue', 'wording', 'cross-story']) {
  const bytes = await fixture(1500);
  const file = `${name}.docx`;
  await writeFile(resolve(root, file), bytes, { flag: 'wx' });
  hashes[file] = createHash('sha256').update(bytes).digest('hex');
}
await writeFile(resolve(root, 'source-hashes.json'), JSON.stringify(hashes, null, 2), { flag: 'wx' });
console.log(root);
