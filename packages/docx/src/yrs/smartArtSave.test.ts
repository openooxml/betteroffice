import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { parseDocx, repackDocxWithWarnings } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { smartArtDrawingRun, smartArtFixture, smartArtNamespaces, smartArtTargets } from './__fixtures__/smartArt';
import { saveEditorDocument, mergeDocxHostMetadata } from './editorSave';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url))));
  await preloadOpcWasm();
}, 30_000);

function makeOpaque(source: ReturnType<typeof smartArtFixture>): void {
  source.parts.set('word/diagrams/drawing1.xml', '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"/>');
  source.parts.set('word/diagrams/layout1.xml', '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:unsupported"/>');
}

for (const [anchored, opaque, linked] of [
  [false, false, false], [true, false, false], [false, true, false], [true, true, false], [false, false, true], [false, true, true],
] as const) {
  for (const seeder of ['native', 'projected'] as const) {
    for (const edit of ['preserve hidden', 'show', 'hide'] as const) {
      it(`${seeder} ${anchored ? 'anchored' : 'inline'} ${opaque ? 'opaque ' : ''}${linked ? 'hyperlinked ' : ''}SmartArt ${edit} survives model serialization`, async () => {
        const source = smartArtFixture(anchored, { hidden: edit !== 'hide' });
        if (opaque) {
          const drawingRun = source.drawingRun.replace('<w:noProof/>', '').replace('<w:lang w:val="en-US"/>', '');
          source.parts.set('word/document.xml', source.parts.get('word/document.xml')!.replace(source.drawingRun, drawingRun));
          source.drawingRun = drawingRun;
          makeOpaque(source);
        }
        if (linked) {
          source.parts.set('word/document.xml', source.parts.get('word/document.xml')!
            .replace(source.drawingRun, `<w:hyperlink w:anchor="diagram">${source.drawingRun}</w:hyperlink>`));
        }
        const bytes = new Uint8Array(rezipPartsToArrayBuffer(new Map(
          [...source.parts].map(([name, xml]) => [name, toBytes(xml)])
        )));
        const host = await parseDocx(bytes.buffer, { preloadFonts: false });
        const session = await createYrsSession();
        const reopened = await createYrsSession();
        try {
          if (seeder === 'native') session.seedFromDocx(bytes);
          else documentToYrs(session, host);
          const embedKind = opaque ? 'opaqueDrawing' : 'shape';
          const segment = session.storySegments('body').find(segment => segment.kind === 'embed' && segment.embedKind === embedKind);
          expect(segment?.kind).toBe('embed');
          if (opaque && segment?.kind === 'embed') expect(segment.payload.kind).toBe('drawingRun');
          expect(segment?.attributes.hidden === true).toBe(edit !== 'hide');
          const { paraId } = session.paragraphs('body')[0]!;
          if (edit !== 'preserve hidden') {
            session.formatRange({ story: 'body', start: { paraId, offset: 24 }, end: { paraId, offset: 25 } }, {
              other: { hidden: edit === 'hide' },
            });
          }
          session.insertText({ story: 'body', paraId, offset: 0 }, 'Edited ');
          const result = await repackDocxWithWarnings(yrsToDocument(session, host));
          const saved = new Uint8Array(result.buffer);
          const parts = unzipContainer(saved);
          const xml = new TextDecoder().decode(parts['word/document.xml']);
          const run = smartArtDrawingRun(xml)!;
          expect(run).toContain(source.drawing);
          if (!opaque) {
            expect(run).toContain('<w:noProof/>');
            expect(run).toContain('<w:lang w:val="en-US"/>');
          }
          expect(run.includes('<w:vanish')).toBe(edit !== 'show');
          if (edit === 'preserve hidden') expect(run).toBe(source.drawingRun);
          expect(xml).toContain('Edited Process overview before ');
          expect(result.warnings.filter(warning => warning.startsWith('Drawing run'))).toEqual([]);
          reopened.openDocx(saved, true);
          const savedSegment = reopened.storySegments('body').find(segment => segment.kind === 'embed' && segment.embedKind === embedKind);
          expect(savedSegment?.kind).toBe('embed');
          expect(savedSegment?.attributes.hidden === true).toBe(edit !== 'show');
        } finally {
          reopened.destroy();
          session.destroy();
        }
      }, 30_000);
    }
  }
}

for (const opaque of [false, true]) {
  for (const seeder of ['native', 'projected'] as const) {
    for (const edit of ['none', 'adjacent', 'hide'] as const) {
      it(`${seeder} ${opaque ? 'opaque ' : ''}SmartArt keeps source position after ${edit}`, async () => {
        const source = smartArtFixture(false);
        const properties = '<w:rPr><w:position w:val="20"/></w:rPr>';
        const drawingRun = source.drawingRun.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, properties);
        source.parts.set('word/document.xml', source.parts.get('word/document.xml')!.replace(source.drawingRun, drawingRun));
        source.drawingRun = drawingRun;
        if (opaque) makeOpaque(source);
        const bytes = new Uint8Array(rezipPartsToArrayBuffer(new Map(
          [...source.parts].map(([name, xml]) => [name, toBytes(xml)])
        )));
        const host = await parseDocx(bytes.buffer, { preloadFonts: false });
        const session = await createYrsSession();
        try {
          if (seeder === 'native') session.seedFromDocx(bytes);
          else documentToYrs(session, host);
          const segment = session.storySegments('body').find(segment => segment.kind === 'embed' &&
            segment.embedKind === (opaque ? 'opaqueDrawing' : 'shape'));
          if (segment?.kind !== 'embed') throw new Error('expected SmartArt');
          if (opaque) expect(segment.payload.kind).toBe('drawingRun');
          expect(segment.attributes.characterSpacing).toBeUndefined();
          const { paraId } = session.paragraphs('body')[0]!;
          if (edit === 'hide') session.formatRange({
            story: 'body', start: { paraId, offset: 24 }, end: { paraId, offset: 25 },
          }, { other: { hidden: true } });
          if (edit !== 'none') session.insertText({ story: 'body', paraId, offset: 0 }, 'Edited ');
          const projected = yrsToDocument(session, host);
          const paragraph = projected.package.document.content[0]!;
          if (paragraph.type !== 'paragraph') throw new Error('expected SmartArt paragraph');
          const run = paragraph.content.find(child => child.type === 'run' && child.content.some(content =>
            content.type === 'shape' || content.type === 'opaqueDrawing'));
          expect(run?.type === 'run' && run.formatting?.position).toBe(20);
          const result = await repackDocxWithWarnings(projected);
          const parts = unzipContainer(new Uint8Array(result.buffer));
          const savedRun = smartArtDrawingRun(new TextDecoder().decode(parts['word/document.xml']))!;
          expect(savedRun).toBe(edit === 'hide'
            ? source.drawingRun.replace('<w:position', '<w:vanish/><w:position')
            : source.drawingRun);
          expect(result.warnings.filter(warning => warning.startsWith('Drawing run'))).toEqual([]);
        } finally {
          session.destroy();
        }
      }, 30_000);
    }
  }
}

for (const seeder of ['native', 'projected'] as const) {
  it(`${seeder} hidden SmartArt in a mixed-content run stays hidden after an adjacent edit`, async () => {
    const source = smartArtFixture(false);
    const mixedRun = `<w:r><w:rPr><w:vanish/></w:rPr><w:t>x</w:t>${source.drawing}</w:r>`;
    source.parts.set('word/document.xml', source.parts.get('word/document.xml')!.replace(source.drawingRun, mixedRun));
    const bytes = new Uint8Array(rezipPartsToArrayBuffer(new Map(
      [...source.parts].map(([name, xml]) => [name, toBytes(xml)])
    )));
    const host = await parseDocx(bytes.buffer, { preloadFonts: false });
    const session = await createYrsSession();
    try {
      if (seeder === 'native') session.seedFromDocx(bytes);
      else documentToYrs(session, host);
      const segment = session.storySegments('body').find(segment => segment.kind === 'embed' && segment.embedKind === 'shape');
      expect(segment?.kind).toBe('embed');
      expect(segment?.attributes.hidden).toBe(true);
      const { paraId } = session.paragraphs('body')[0]!;
      session.insertText({ story: 'body', paraId, offset: 0 }, 'Edited ');
      const result = await repackDocxWithWarnings(yrsToDocument(session, host));
      const xml = new TextDecoder().decode(unzipContainer(new Uint8Array(result.buffer))['word/document.xml']);
      const run = smartArtDrawingRun(xml)!;
      expect(run).toContain('<dgm:relIds');
      expect(run).toContain('<w:vanish');
      expect(xml).toContain('Edited Process overview before ');
    } finally {
      session.destroy();
    }
  }, 30_000);
}

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
