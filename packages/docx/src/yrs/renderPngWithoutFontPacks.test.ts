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

for (const { kind } of ENGINES) {
  describe.skipIf(engineKind !== kind)(`renderPng without the optional font packs on the ${kind} engine`, () => {
    test('faces whose pack is not installed are skipped and the export succeeds', async () => {
      stubBrowser();
      const session = await createYrsSession();
      try {
        session.openDocx(scriptDocument(), true);
        expect(await session.renderPng({ page: 0 })).toBeInstanceOf(Uint8Array);
        expect(importLatin).toHaveBeenCalled();
      } finally { session.destroy(); }
    });
  });
}
