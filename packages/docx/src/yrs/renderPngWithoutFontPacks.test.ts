import { afterEach, beforeAll, describe, expect, mock, test } from 'bun:test';
import { engineKind } from '../wasm/enginePng';
import { createYrsSession } from './index';
import { ENGINES, preloadEngines, restoreBrowser, scriptDocument, stubBrowser } from './renderPngTestSupport';

const importLatin = mock(() => {
  throw new Error("Cannot find package '@betteroffice/fonts-latin'");
});
mock.module('@betteroffice/fonts-latin', () => ({
  then(resolve: (addon: unknown) => void, reject: (error: unknown) => void) {
    Promise.resolve().then(importLatin).then(resolve, reject);
  },
}));
mock.module('@betteroffice/fonts-intl', () => ({
  then(_resolve: unknown, reject: (error: unknown) => void) {
    Promise.resolve().then(() => reject(new Error("Cannot find package '@betteroffice/fonts-intl'")));
  },
}));

beforeAll(preloadEngines);
afterEach(restoreBrowser);

async function exported(): Promise<unknown> {
  const session = await createYrsSession();
  try {
    session.openDocx(scriptDocument(), true);
    return await session.renderPng({ page: 0 }).catch((error: unknown) => error);
  } finally { session.destroy(); }
}

for (const { kind } of ENGINES) {
  describe.skipIf(engineKind !== kind)(`renderPng without the optional font packs on the ${kind} engine`, () => {
    if (kind === 'wasm') {
      test('with no Latin pack installed the wasm engine draws with the browser\'s fonts, as before, while the native engine refuses the export', async () => {
        stubBrowser();
        expect(await exported()).toBeInstanceOf(Uint8Array);
        expect(importLatin).toHaveBeenCalled();
      });
    } else {
      test('with no Latin pack installed the native engine refuses the export naming the family, while the wasm engine draws with the browser\'s fonts, as before', async () => {
        const error = await exported();
        expect(importLatin).toHaveBeenCalled();
        expect(error).toMatchObject({ name: 'PngExportError', code: 'font-unavailable' });
        expect((error as Error).message).toContain('missing font chain for `calibri|0|0`');
      });
    }
  });
}
