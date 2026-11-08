import { expect, test } from 'bun:test';
import intlSizes from './intl-sizes.json';
import { checkInstalledPacks } from './packDiagnostics';

const remote = (files: string[]) => Object.fromEntries(files.map((file) => [file, () => new URL(`https://cdn.example/intl/${file}`)]));
const intlFiles = Object.keys(intlSizes);
const none = async () => undefined;

test('a pack that is not installed is not reported', async () => {
  expect(await checkInstalledPacks({ latin: none, intl: none, cjk: none })).toEqual([]);
});

test('in a browser, an installed pack listing every file of this release is not reported', async () => {
  expect(await checkInstalledPacks({ latin: none, intl: async () => remote(intlFiles), cjk: none })).toEqual([]);
});

test('in a browser, an installed pack missing files of this release is reported as another release', async () => {
  const [diagnostic] = await checkInstalledPacks({ latin: none, intl: async () => remote(intlFiles.slice(1)), cjk: none });
  expect(diagnostic).toMatchObject({ pack: 'intl', reason: 'version-mismatch' });
  expect(diagnostic!.detail).toStartWith('@betteroffice/fonts-intl of unknown version does not match the ');
  expect(diagnostic!.detail).toEndWith('; its faces are substituted');
});

test('an installed pack whose asset list cannot be read is reported and the other packs are still checked', async () => {
  const diagnostics = await checkInstalledPacks({
    latin: none,
    intl: async () => { throw new TypeError('Invalid INTL_FONT_ASSET_URLS export from @betteroffice/fonts-intl'); },
    cjk: async () => remote([]),
  });
  expect(diagnostics.map(({ pack, reason }) => [pack, reason])).toEqual([['intl', 'invalid-manifest'], ['cjk', 'version-mismatch']]);
});
