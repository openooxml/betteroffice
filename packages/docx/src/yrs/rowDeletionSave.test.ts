import { describe, expect, it } from 'bun:test';
import { SAVE_PATHS, SOURCE_PARAGRAPHS, SOURCE_ROWS, TABLE_PREFIX, W, W14, setupStructuralSaves, structuralFixture } from './structuralSaveTestSupport';

const save = setupStructuralSaves();

describe.each([...SAVE_PATHS])('%s row deletion preserves source bytes', (path) => {
  describe.each(['plain', 'w14'])('with a %s root', (binding) => {
    const root = `<w:document xmlns:w="${W}"${binding === 'w14' ? ` xmlns:w14="${W14}"` : ''}>`;
    const source = structuralFixture(root);
    it.each([0, 1])('drops row %i and keeps the other row and paragraphs', async (row) => {
      const xml = await save(path, source, (engine) => {
        const at = { story: 'body', tableIndex: 0, row, column: 0 };
        engine.deleteRow({ anchor: at, head: at });
      });
      for (const paragraph of SOURCE_PARAGRAPHS) expect(xml).toContain(paragraph);
      expect(xml).toContain(TABLE_PREFIX);
      expect(xml).toContain(SOURCE_ROWS[1 - row]!);
      expect(xml).not.toContain(row === 0 ? 'first row' : 'last row');
      expect(xml.match(/<w:tr>/g)).toHaveLength(1);
      expect(xml.startsWith(root)).toBe(true);
    });
  });
});
