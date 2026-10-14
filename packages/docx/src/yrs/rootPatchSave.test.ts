import { describe, expect, it } from 'bun:test';
import { MC, SAVE_PATHS, SOURCE_PARAGRAPHS, W, W14, insertStructuralParagraph, setupStructuralSaves, structuralFixture } from './structuralSaveTestSupport';

const save = setupStructuralSaves();
const PROLOG = "<?xml version='1.0' encoding='UTF-8'?>\n";
const ROOTS = [
  `<w:document  xmlns:x='urn:source' x:token = 'a&amp;b' xmlns:w = '${W}' >`,
  `<w:document xmlns:x='urn:source' xmlns:mc='${MC}' mc:Ignorable = 'x' xmlns:w='${W}' x:token='a&amp;b' >`,
  `<w:document xmlns:w='${W}' xmlns:mc='${MC}' xmlns:x='urn:source' x:token='a&amp;b' >`,
];

describe.each([...SAVE_PATHS])('%s paragraph insertion patches the source root', (path) => {
  it.each(ROOTS)('preserves every other root byte and declaration order: %s', async (root) => {
    const xml = await save(path, structuralFixture(PROLOG + root), insertStructuralParagraph);
    expect(xml.startsWith(PROLOG)).toBe(true);
    const writtenRoot = xml.match(/<w:document\b[^>]*>/)![0];
    expect(writtenRoot).toContain(`xmlns:w14="${W14}"`);
    expect(writtenRoot.match(/mc:Ignorable\s*=\s*['"]([^'"]*)['"]/)![1]!.split(/\s+/)).toContain('w14');
    const restored = writtenRoot
      .replace(` xmlns:w14="${W14}"`, '')
      .replace(root.includes('xmlns:mc') ? /$^/ : new RegExp(` xmlns:mc="${MC}"`), '')
      .replace(root.includes('mc:Ignorable') ? 'x w14' : ' mc:Ignorable="w14"', root.includes('mc:Ignorable') ? 'x' : '');
    expect(restored).toBe(root);
    for (const paragraph of SOURCE_PARAGRAPHS) expect(xml).toContain(paragraph);
    expect(xml).toContain('>inserted</w:t>');
  });
});
