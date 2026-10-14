import { describe, expect, it } from 'bun:test';
import { MC, SAVE_PATHS, W, W14, setupStructuralSaves, structuralFixture } from './structuralSaveTestSupport';

const save = setupStructuralSaves();
const ROOT = `<w:document xmlns:x='urn:source' x:token = 'a&amp;b' xmlns:w='${W}' xmlns:w14='${W14}' xmlns:mc='${MC}' mc:Ignorable='w14' >`;

describe.each([...SAVE_PATHS])('%s whole-part fallback keeps the source root', (path) => {
  it.each(['', "<?xml version='1.0' encoding='UTF-8'?>\n"])('preserves the source prolog %j on a column insertion', async (prolog) => {
    const xml = await save(path, structuralFixture(prolog + ROOT), (engine) => {
      engine.insertColumn({ story: 'body', tableIndex: 0, row: 0, column: 0 }, 'right');
    });
    expect(xml.startsWith(prolog + ROOT)).toBe(true);
    expect(xml.match(/<\?xml\b/g)?.length ?? 0).toBe(prolog === '' ? 0 : 1);
    expect(xml.match(/<w:gridCol\b/g)).toHaveLength(2);
    expect(xml.match(/<w:tc>/g)).toHaveLength(4);
  });
});
