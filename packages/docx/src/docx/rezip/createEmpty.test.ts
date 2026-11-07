import { expect, test } from 'bun:test';

import { parseDocx } from '../../docx';
import { repackDocx } from '../rezip';
import { createEmptyDocx } from './createEmpty';

function normalTable(document: Awaited<ReturnType<typeof parseDocx>>) {
  return document.package.styles?.styles?.find(
    (style) => style.type === 'table' && style.default === true
  );
}

function expectWordNormalTable(style: ReturnType<typeof normalTable>) {
  expect(style?.styleId).toBe('TableNormal');
  expect(style?.name).toBe('Normal Table');
  expect(style?.tblPr?.cellMargins).toEqual({
    top: { value: 0, type: 'dxa' },
    left: { value: 108, type: 'dxa' },
    bottom: { value: 0, type: 'dxa' },
    right: { value: 108, type: 'dxa' },
  });
}

test('a new blank document carries Word\'s Normal Table style as its default table style', async () => {
  const document = await parseDocx(await createEmptyDocx(), { preloadFonts: false });
  expectWordNormalTable(normalTable(document));
});

test('saving a new blank document keeps its Normal Table style', async () => {
  const document = await parseDocx(await createEmptyDocx(), { preloadFonts: false });
  const saved = await repackDocx(document);
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expectWordNormalTable(normalTable(reopened));
});

test('a table inserted into a new blank document gets Word\'s 0.19 cm left and right cell margins', async () => {
  const { readFileSync } = await import('node:fs');
  const { resolve } = await import('node:path');
  const { preloadEditWasm } = await import('../../wasm/edit');
  const { createYrsSession } = await import('../../yrs');
  const { buildResidentRegionLayoutRequest } = await import('../../editor/computeLayout');
  await preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../../wasm/generated/edit/docx_edit_bg.wasm')))
  );
  const bytes = new Uint8Array(await createEmptyDocx());
  const document = await parseDocx(bytes.slice().buffer, { preloadFonts: false });
  const hostEnv = buildResidentRegionLayoutRequest(document, 24, {}).renderEnv;
  const session = await createYrsSession({ clientId: 23121 });
  try {
    session.seedFromDocx(bytes);
    const anchor = session.paragraphs('body')[0]!;
    session.insertTable({ story: 'body', paraId: anchor.paraId, offset: 0 }, 1, 2);
    for (const env of [{}, hostEnv]) {
      const blocks = session.yrsBlocksForStory('body', env) as Array<{
        kind: string;
        rows?: Array<{ cells: Array<{ padding?: { left: number; right: number } }> }>;
      }>;
      const padding = blocks.find((block) => block.kind === 'table')?.rows?.[0]?.cells[0]?.padding;
      expect(padding?.left).toBeCloseTo(7.2, 6);
      expect(padding?.right).toBeCloseTo(7.2, 6);
    }
  } finally {
    session.destroy();
  }
}, 30_000);
