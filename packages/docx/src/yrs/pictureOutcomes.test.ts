import { beforeAll, expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { engineKind } from '../wasm/enginePng';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { createYrsSession } from './index';
import { PICTURE_CASES, pictureDocument, pictureOutcomeFailures, type PictureResult, type ReportedPicture } from '../../../../e2e/docx-png-pictures';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
});

// The wasm engine exports through a canvas, which this runtime lacks; e2e/browser/docx-png-pictures.browser.ts runs the same table there.
test.skipIf(engineKind !== 'native')('native PNG export follows the picture outcome table shared with the canvas engine', async () => {
  const results: Record<string, PictureResult> = {};
  for (const { name, part, strictPictures } of PICTURE_CASES) {
    const session = await createYrsSession();
    const pictures: ReportedPicture[] = [];
    try {
      session.openDocx(await pictureDocument(part), true);
      const png = await session.renderPng({ page: 0 }, {
        strictPictures,
        onDiagnostic: (diagnostic) => pictures.push(...diagnostic.pictures.map(({ id, reason }) => ({ id, reason }))),
      });
      results[name] = { png: createHash('sha256').update(png).digest('hex'), pictures };
    } catch (error) {
      results[name] = { code: (error as { code?: string }).code, message: String((error as Error).message), pictures };
    } finally { session.destroy(); }
  }
  expect(pictureOutcomeFailures(results)).toEqual([]);
});
