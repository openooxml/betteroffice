import { expect, test } from 'bun:test';
import { readResidentToolbarMetadata } from './residentToolbarMetadata';

test('toolbar_metadata_keeps_custom_style_previews_without_document_or_engine_peers', () => {
  const metadata = readResidentToolbarMetadata(JSON.stringify({ referencedFonts: ['Zed', 'Ada', 'Ada'], envelope: { wireVersion: 1, document: { package: { styles: { styles: [
    { styleId: 'Hidden', type: 'paragraph', hidden: true },
    { styleId: 'Character', type: 'character' },
    { styleId: 'Custom', type: 'paragraph', name: 'Custom heading', uiPriority: 1, rPr: { fontSize: 32, bold: true, color: { rgb: '123456' } } },
    { styleId: 'Normal', type: 'paragraph', name: 'Normal', uiPriority: 0 },
  ] } } } } }));
  expect(metadata).toEqual({ styles: [
    { styleId: 'Normal', name: 'Normal', priority: 0, fontSize: undefined, bold: undefined, italic: undefined, color: undefined },
    { styleId: 'Custom', name: 'Custom heading', priority: 1, fontSize: 32, bold: true, italic: undefined, color: '123456' },
  ], fontFamilies: ['Ada', 'Zed'] });
});
