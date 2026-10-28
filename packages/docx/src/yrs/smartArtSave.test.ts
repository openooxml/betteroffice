import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { smartArtDrawingRun, smartArtFixture, smartArtNamespaces, smartArtTargets } from './__fixtures__/smartArt';
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

for (const anchored of [false, true]) {
  for (const scenario of ['resize', 'body namespace', 'large extension'] as const) {
    it(`native editor saves ${anchored ? 'anchored' : 'inline'} SmartArt with ${scenario}`, async () => {
      const extensionSize = scenario === 'large extension' ? 1_100_000 : undefined;
      const source = smartArtFixture(anchored, { bodyNamespace: scenario === 'body namespace', extensionSize });
      const bytes = new Uint8Array(rezipPartsToArrayBuffer(new Map(
        [...source.parts].map(([name, xml]) => [name, toBytes(xml)])
      )));
      const session = await createYrsSession();
      try {
        const host = session.openDocx(bytes, true).document;
        const segment = session.storySegments('body').find(segment => segment.kind === 'embed' && segment.embedKind === 'shape');
        if (segment?.kind !== 'embed') throw new Error('expected SmartArt');
        const shape = JSON.parse(segment.payload.shapeJson as string);
        if (extensionSize) {
          expect(shape.drawingXml).toBeUndefined();
          expect(shape.drawingRunXml).toContain('x'.repeat(extensionSize));
          expect((segment.payload.shapeJson as string).length).toBeLessThan(2_000_000);
        }
        if (scenario === 'resize') {
          shape.size.width = 2743200;
          session.applyRawOps('body', [{ op: 'setEmbedAttr', index: 24, key: 'shapeJson', value: JSON.stringify(shape) }]);
        } else {
          const { paraId } = session.paragraphs('body')[0]!;
          session.insertText({ story: 'body', paraId, offset: 0 }, 'Edited ');
          session.setParagraphAttrs({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
            { defaultTextFormatting: { fontSize: 40 } });
        }
        const base = session.materializeDocx()!;
        const saved = new Uint8Array(await saveEditorDocument(session,
          yrsToDocument(session, mergeDocxHostMetadata(base, host)), []));
        const parts = unzipContainer(saved);
        const xml = new TextDecoder().decode(parts['word/document.xml']);
        const relationships = new TextDecoder().decode(parts['word/_rels/document.xml.rels']);
        expect(smartArtTargets(xml, relationships)).toHaveLength(4);
        expect(smartArtNamespaces(xml)).toEqual(['http://schemas.openxmlformats.org/drawingml/2006/diagram']);
        expect(xml).toContain(`cx="${scenario === 'resize' ? 2743200 : 5486400}" cy="3200400"`);
        expect(xml).not.toContain('<wps:wsp');
        if (extensionSize) expect(xml).toContain('x'.repeat(extensionSize));
      } finally { session.destroy(); }
    }, 30_000);
  }
}
