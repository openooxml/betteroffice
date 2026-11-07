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
