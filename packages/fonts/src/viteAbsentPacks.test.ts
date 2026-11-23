import { expect, mock, test } from 'bun:test';
import { resolveScriptFallbackFace } from './manifest';

let missingMessage = (pack: string) => `Cannot find package '@betteroffice/fonts-${pack}'`;
for (const pack of ['intl', 'cjk']) {
  mock.module(`@betteroffice/fonts-${pack}`, () => ({
    then(_resolve: unknown, reject: (error: unknown) => void) {
      Promise.resolve().then(() => reject(new Error(missingMessage(pack))));
    },
  }));
}

test('Vite 7 and 8 stubs for the optional packs count as absent packs', async () => {
  const { fontAssetUrls } = await import('./index');
  for (const suffix of ['', ' Is it installed?']) {
    missingMessage = (pack) => `Could not resolve "@betteroffice/fonts-${pack}" imported by "@betteroffice/fonts".${suffix}`;
    for (const pack of ['intl', 'cjk'] as const) {
      await expect(fontAssetUrls(pack)).resolves.toBeUndefined();
    }
  }
});

test('other import failures of the optional packs still throw', async () => {
  const { fontAssetUrls } = await import('./index');
  for (const pack of ['intl', 'cjk'] as const) {
    const failure = `Could not resolve "@betteroffice/fonts-${pack}" imported by "another-package".`;
    missingMessage = () => failure;
    await expect(fontAssetUrls(pack)).rejects.toThrow(failure);
  }
  missingMessage = () => 'Unexpected token in fonts-intl';
  await expect(fontAssetUrls('intl')).rejects.toThrow('Unexpected token in fonts-intl');
});

test('a host loading an intl face under a Vite stub gets a missing add-on error', async () => {
  const { isMissingFontAddon, loadBundledFontBytes } = await import('./index');
  missingMessage = (pack) => `Could not resolve "@betteroffice/fonts-${pack}" imported by "@betteroffice/fonts".`;
  const error = await loadBundledFontBytes(resolveScriptFallbackFace('thai', false, false)!).then(
    () => undefined, (reason: unknown) => reason);
  expect(isMissingFontAddon(error)).toBe(true);
});
