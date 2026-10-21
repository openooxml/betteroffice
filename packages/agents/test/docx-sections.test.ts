import { expect, test } from 'bun:test';
import { openDocx } from '../src';
import { docxBatchSchema } from '../src/docx-schema';
import { sectionPackage } from '../../docx/src/yrs/sectionTestPackage';

test('agents insert before a paragraph, set a versioned section and reopen exported sections', async () => {
  const doc = await openDocx(await sectionPackage());
  try {
    const outline = doc.list();
    await doc.edit({ version: outline.version, edits: [{ op: 'insert_section_break', before: outline.items[1].ref, kind: 'nextPage' }] });
    const fresh = doc.list({ sectionOffset: 1, sectionLimit: 1 });
    expect(fresh.sections.items[0]).toMatchObject({ number: 2, firstParagraph: outline.items[1].ref, lastParagraph: outline.items[2].ref });
    await doc.edit({ version: fresh.version, edits: [{ op: 'set_section', section: 2, orientation: 'landscape', margins: { top: 900, bottom: 900, left: 1000, right: 1000 } }] });
    await expect(doc.edit({ version: outline.version, edits: [{ op: 'set_section', section: 1, orientation: 'portrait' }] })).rejects.toMatchObject({ code: 'STALE_VERSION' });
    const reopened = await openDocx(await doc.export());
    try { expect(reopened.list().sections.items[1]).toMatchObject({ number: 2, orientation: 'landscape', pageWidth: 15840, pageHeight: 12240, margins: { top: 900, left: 1000 } }); }
    finally { reopened.close(); }
    const before = doc.session.encodeState();
    await expect(doc.edit({ version: doc.list().version, edits: [
      { op: 'set_section', section: 1, orientation: 'landscape' },
      { op: 'insert_section_break', before: outline.items[0].ref, kind: 'continuous' },
    ] })).rejects.toMatchObject({ code: 'UNSUPPORTED_TARGET' });
    expect(doc.session.encodeState()).toEqual(before);
  } finally { doc.close(); }
});

test('section break schema requires exactly one before or after ref', () => {
  const batch = (edit: unknown) => docxBatchSchema.safeParse({ version: 0, edits: [edit] }).success;
  expect(batch({ op: 'insert_section_break', before: 'p1', kind: 'oddPage' })).toBe(true);
  expect(batch({ op: 'insert_section_break', after: 'p1', kind: 'evenPage' })).toBe(true);
  expect(batch({ op: 'insert_section_break', kind: 'continuous' })).toBe(false);
  expect(batch({ op: 'insert_section_break', before: 'p1', after: 'p2', kind: 'continuous' })).toBe(false);
});
