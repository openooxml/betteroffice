import { mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { renderPages, type SiteData } from './render';
import { newestFirst, type Suite } from './stats';
import { validateSuite } from '../../../eval/formats/collect-results';

export async function loadData(): Promise<SiteData> {
  const directory = resolve(import.meta.dir, '../../../eval/results');
  const suites: Suite[] = [];
  for (const name of (await readdir(directory)).sort()) {
    if (!name.endsWith('.json') || name.endsWith('.reasons.json') || name.endsWith('.meta.json')) continue;
    const suite = JSON.parse(await readFile(resolve(directory, name), 'utf8')) as Suite;
    validateSuite(suite);
    if (suites.some(s => s.suite === suite.suite)) throw new Error(`Duplicate suite: ${suite.suite}`);
    suites.push(suite);
  }
  return {
    suites: newestFirst(suites),
    frame: await readFile(resolve(import.meta.dir, 'frame.html'), 'utf8'),
  };
}

export async function buildSite(directory = resolve(import.meta.dir, '../dist')) {
  const data = await loadData();
  const pages = renderPages(data);
  await rm(directory, { recursive: true, force: true });
  await mkdir(resolve(directory, 'results'), { recursive: true });
  for (const [name, html] of pages) {
    const target = resolve(directory, name);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, html);
  }
  for (const suite of data.suites) await writeFile(resolve(directory, 'results', encodeURIComponent(suite.suite) + '.json'), JSON.stringify(suite, null, 2) + '\n');
  await writeFile(resolve(directory, 'icon.svg'), await readFile(resolve(import.meta.dir, 'icon.svg')));
  return pages;
}

if (import.meta.main) await buildSite();
