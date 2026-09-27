import { mkdir, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pptxFixture, xlsxFixture } from '../test/format-fixtures';

const root = resolve(process.argv[2] ?? '/tmp/betteroffice-formats-eval');
for (const task of ['xlsx', 'pptx', 'mixed']) {
  const directory = resolve(root, task);
  await mkdir(directory, { recursive: true });
  if (task !== 'pptx') await writeFile(resolve(directory, 'budget.xlsx'), await xlsxFixture(), { flag: 'wx' });
  if (task !== 'xlsx') await writeFile(resolve(directory, 'slides.pptx'), await pptxFixture(), { flag: 'wx' });
}
console.log(root);
