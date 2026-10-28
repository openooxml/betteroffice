import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { smartArtDrawingRun, smartArtFixture, smartArtTargets } from './__fixtures__/smartArt';
import { saveEditorDocument, mergeDocxHostMetadata } from './editorSave';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url))));
  await preloadOpcWasm();
});

for (const anchored of [false, true]) {
  it(`saveEditorDocument preserves ${anchored ? 'anchored' : 'inline'} SmartArt in an edited paragraph`, async () => {
    const source = smartArtFixture(anchored);
    const bytes = new Uint8Array(rezipPartsToArrayBuffer(new Map(
      [...source.parts].map(([name, xml]) => [name, toBytes(xml)])
    )));
    const session = await createYrsSession();
    const reopened = await createYrsSession();
    try {
      const host = session.openDocx(bytes, true).document;
      const paragraph = session.paragraphs('body')[0]!;
      const offset = paragraph.text.indexOf('before');
      session.replaceRange({
        story: 'body', start: { paraId: paragraph.paraId, offset },
        end: { paraId: paragraph.paraId, offset: offset + 'before'.length },
      }, 'ahead of');
      const base = session.materializeDocx()!;
      const projected = yrsToDocument(session, mergeDocxHostMetadata(base, host));
      const saved = new Uint8Array(await saveEditorDocument(session, projected, []));
      const parts = unzipContainer(saved);
      const text = (name: string) => new TextDecoder().decode(parts[name]);
      const xml = text('word/document.xml');
      expect(smartArtDrawingRun(xml)).toBe(source.drawingRun);
      expect(xml).toContain('Process overview ahead of ');
      const targets = smartArtTargets(xml, text('word/_rels/document.xml.rels'));
      expect(targets).toHaveLength(4);
      for (const target of targets) expect(parts[target]).toBeDefined();
      const drawingTarget = text('word/diagrams/_rels/data1.xml.rels').match(/Target="([^"]+)"/)![1]!;
      expect(parts[`word/diagrams/${drawingTarget}`]).toBeDefined();
      for (const [name, original] of source.parts) {
        if (name !== 'word/document.xml') expect(text(name)).toBe(original);
      }
      const plan = session.paragraphSavePlan().splicedParts!.find(({ part }) => part === 'word/document.xml')!;
      expect(plan.legacyParagraphs ?? []).toEqual([]);
      expect(plan.paragraphPatches).toHaveLength(1);
      reopened.openDocx(saved, true);
      expect(reopened.paragraphs('body')[0]?.text).toContain('Process overview ahead of ');
      expect(reopened.readParagraphs({ story: 'body', view: 'accepted' })).toMatchObject({
        ok: true, paragraphs: [expect.objectContaining({ atoms: [expect.objectContaining({ kind: 'other' })] }), expect.anything()],
      });
    } finally {
      reopened.destroy();
      session.destroy();
    }
  }, 30_000);
}
