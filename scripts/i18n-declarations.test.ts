import { expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const root = resolve(import.meta.dir, '..');

for (const format of ['docx', 'xlsx', 'pptx', 'vsdx']) {
  test(`${format} locale exports typecheck in an isolated consumer`, () => {
    const directory = mkdtempSync(join(tmpdir(), 'betteroffice-i18n-types-'));
    try {
      const source = join(root, 'packages', `${format}-i18n`);
      const manifest = JSON.parse(readFileSync(join(source, 'package.json'), 'utf8'));
      const destination = join(directory, 'node_modules', manifest.name);
      mkdirSync(destination, { recursive: true });
      writeFileSync(join(destination, 'package.json'), JSON.stringify(manifest));

      const build = spawnSync(process.execPath, [
        'run', 'build', '--out-dir', join(destination, 'dist'),
      ], { cwd: source, encoding: 'utf8' });
      if (build.status !== 0) throw new Error(build.stdout + build.stderr);

      const imports = Object.keys(manifest.exports)
        .filter((entry) => entry !== './package.json')
        .map((entry, index) => {
          const specifier = manifest.name + (entry === '.' ? '' : entry.slice(1));
          return `import * as locale${index} from '${specifier}'; void locale${index};`;
        });
      writeFileSync(join(directory, 'consumer.mts'), `${imports.join('\n')}
import { en, createT, type LocaleStrings, type TranslationKey, type Translations } from '${manifest.name}';
const strings: LocaleStrings = { ...en, _lang: 'custom' };
const translations: Translations = { _lang: 'custom' };
const key: TranslationKey = '_lang';
const result: string = createT(strings)(key);
void [translations, result];
type Assert<T extends true> = T;
export type UnknownKeysRejected = Assert<'invalid.translation.key' extends TranslationKey ? false : true>;
`);
      writeFileSync(join(directory, 'tsconfig.json'), JSON.stringify({
        compilerOptions: {
          target: 'ES2022',
          module: 'NodeNext',
          moduleResolution: 'NodeNext',
          strict: true,
          skipLibCheck: false,
          noEmit: true,
          types: [],
        },
        files: ['consumer.mts'],
      }));
      const tsc = createRequire(join(source, 'package.json')).resolve('typescript/bin/tsc');
      const check = spawnSync(process.execPath, [
        tsc, '-p', directory,
      ], { cwd: directory, encoding: 'utf8' });
      expect({ status: check.status, output: check.stdout + check.stderr }).toEqual({ status: 0, output: '' });
    } finally {
      rmSync(directory, { recursive: true, force: true });
    }
  }, 60_000);
}
