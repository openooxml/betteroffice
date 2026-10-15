import { mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { renderPages, type SiteData } from './render';
import { newestFirst, type Suite } from './stats';

export async function loadData(): Promise<SiteData> {
  const directory = resolve(import.meta.dir, '../../../eval/results');
  const suites: Suite[] = [];
  for (const name of (await readdir(directory)).sort()) {
    if (!name.endsWith('.json') || name.endsWith('.reasons.json')) continue;
    const suite = JSON.parse(await readFile(resolve(directory, name), 'utf8')) as Suite;
    if (suite.schema !== 1 || !suite.suite || !suite.runs.length) throw new Error(`Invalid suite: ${name}`);
    if (suites.some(s => s.suite === suite.suite)) throw new Error(`Duplicate suite: ${suite.suite}`);
    suites.push(suite);
  }
  const readme = await readFile(resolve(import.meta.dir, '../../../eval/README.md'), 'utf8');
  const running = readme.split('### Running')[1]?.split('## Usability tasks')[0];
  const commands = [...(running ?? '').matchAll(/```sh\n([\s\S]*?)```/g)].map(match => match[1]!.trim());
  if (commands.length !== 3) throw new Error('Expected build/run, generic, and checker commands in eval/README.md');
  return {
    suites: newestFirst(suites), commands,
    fixtures: JSON.parse(await readFile(resolve(import.meta.dir, '../../../eval/formats/fixtures.json'), 'utf8')),
    tasks: JSON.parse(await readFile(resolve(import.meta.dir, '../../../eval/formats/tasks.json'), 'utf8')),
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
