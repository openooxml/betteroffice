import { expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx } from '../src';
import { saveEditorDocument, mergeDocxHostMetadata, yrsToDocument } from '@betteroffice/docx/yrs';
import { smartArtDrawingRun, smartArtFixture, smartArtNamespaces, smartArtTargets } from '../../docx/src/yrs/__fixtures__/smartArt';

for (const anchored of [false, true]) {
  for (const path of ['agent export', 'editor save'] as const) {
    for (const fallback of [false, true]) {
      test(`${path} keeps ${anchored ? 'anchored' : 'inline'} SmartArt beside edited text${fallback ? ' through model serialization' : ''}`, async () => {
        const source = smartArtFixture(anchored);
        const zip = new JSZip();
        for (const [name, xml] of source.parts) zip.file(name, xml);
        const doc = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
        try {
          const match = doc.grep({ query: 'before' }).matches[0]!;
          const proposal = doc.propose({ author: 'test', edits: [{ match: match.match, newText: 'ahead of' }] });
          expect(await doc.accept(proposal.id, { tracked: false })).toMatchObject({ status: 'accepted' });
          if (fallback) {
            const paragraph = doc.session.paragraphs('body')[0]!;
            doc.session.setParagraphAttrs({
              story: 'body', start: { paraId: paragraph.paraId, offset: 0 },
              end: { paraId: paragraph.paraId, offset: 0 },
            }, { defaultTextFormatting: { fontSize: 40 } });
          }
          const plan = doc.session.paragraphSavePlan().splicedParts!.find(({ part }) => part === 'word/document.xml')!;
          let bytes: Uint8Array;
          if (path === 'editor save') {
            const base = doc.session.materializeDocx()!;
            bytes = new Uint8Array(await saveEditorDocument(
              doc.session, yrsToDocument(doc.session, mergeDocxHostMetadata(base, base)), []
            ));
          } else bytes = await doc.export();
          const saved = await JSZip.loadAsync(bytes);
          const xml = await saved.file('word/document.xml')!.async('string');
          expect(smartArtDrawingRun(xml)).toBe(source.drawingRun);
          expect(xml).toContain('Process overview ahead of ');
          expect(xml).not.toContain('<wps:wsp');
          for (const [name, original] of source.parts) {
            if (name !== 'word/document.xml') expect(await saved.file(name)!.async('string')).toBe(original);
          }
          const relationships = await saved.file('word/_rels/document.xml.rels')!.async('string');
          const targets = smartArtTargets(xml, relationships);
          expect(targets).toHaveLength(4);
          for (const target of targets) expect(saved.file(target)).not.toBeNull();
          const dataRelationships = await saved.file('word/diagrams/_rels/data1.xml.rels')!.async('string');
          expect(dataRelationships).toContain('relationships/diagramDrawing');
          expect(dataRelationships).toContain('Id="rId14"');
          const target = dataRelationships.match(/Target="([^"]+)"/)![1]!;
          expect(saved.file(`word/diagrams/${target}`)).not.toBeNull();
          expect(await saved.file('word/diagrams/data1.xml')!.async('string')).toContain('relId="rId14"');
          if (fallback) expect(plan.legacyParagraphs?.length).toBeGreaterThan(0);
          else {
            expect(plan.legacyParagraphs ?? []).toEqual([]);
            expect(plan.paragraphPatches).toHaveLength(1);
          }
          const reopened = await openDocx(bytes);
          try {
            expect(reopened.grep({ query: 'ahead of' }).matches).toHaveLength(1);
            expect(reopened.session.readParagraphs({ story: 'body', view: 'accepted' }))
              .toMatchObject({ ok: true, paragraphs: [expect.objectContaining({ atoms: [expect.objectContaining({ kind: 'other' })] }), expect.anything()] });
          } finally { reopened.close(); }
        } finally { doc.close(); }
      }, 30_000);
    }
  }
}

for (const anchored of [false, true]) {
  for (const editor of [false, true]) {
    for (const scenario of ['resize', 'body namespace', 'large extension', 'oversized extension'] as const) {
      test(`${editor ? 'editor' : 'agent'} save keeps ${anchored ? 'anchored' : 'inline'} SmartArt with ${scenario}`, async () => {
        const extensionSize = scenario === 'large extension' ? 1_100_000 : scenario === 'oversized extension' ? 2_100_000 : undefined;
        const source = smartArtFixture(anchored, { bodyNamespace: scenario === 'body namespace', extensionSize });
        const zip = new JSZip();
        for (const [name, xml] of source.parts) zip.file(name, xml);
        const doc = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
        try {
          const segment = doc.session.storySegments('body').find(segment => segment.kind === 'embed' && segment.embedKind === 'shape');
          if (segment?.kind !== 'embed') throw new Error('expected SmartArt shape');
          const json = segment.payload.shapeJson as string;
          const shape = JSON.parse(json) as { size: { width: number }; drawingXml?: string; drawingRunXml?: string };
          if (extensionSize) {
            expect(shape.drawingXml).toBeUndefined();
            expect(shape.drawingRunXml).toContain('x'.repeat(extensionSize));
            expect(json.match(/<keep:payload/g)).toHaveLength(1);
            if (extensionSize === 1_100_000) expect(json.length).toBeLessThan(2_000_000);
            else expect(json.length).toBeGreaterThan(2_000_000);
          }
          if (scenario === 'resize') {
            shape.size.width = 2_743_200;
            doc.session.applyRawOps('body', [{ op: 'setEmbedAttr', index: 'Process overview before '.length, key: 'shapeJson', value: JSON.stringify(shape) }]);
          } else {
            const match = doc.grep({ query: 'before' }).matches[0]!;
            const proposal = doc.propose({ author: 'test', edits: [{ match: match.match, newText: 'ahead of' }] });
            await doc.accept(proposal.id, { tracked: false });
            const { paraId } = doc.session.paragraphs('body')[0]!;
            doc.session.setParagraphAttrs({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
              { defaultTextFormatting: { fontSize: 40 } });
          }
          const base = doc.session.materializeDocx()!;
          const bytes = editor ? new Uint8Array(await saveEditorDocument(doc.session,
            yrsToDocument(doc.session, mergeDocxHostMetadata(base, base)), [])) : await doc.export();
          const saved = await JSZip.loadAsync(bytes);
          const xml = await saved.file('word/document.xml')!.async('string');
          expect(xml).toContain(`cx="${scenario === 'resize' ? 2743200 : 5486400}" cy="3200400"`);
          expect(xml).not.toContain('<wps:wsp');
          const run = smartArtDrawingRun(xml)!;
          expect(run.replace(/\sxmlns(?::[\w.-]+)?="[^"]*"/g, '')).toBe(
            source.drawingRun.replace(/\sxmlns(?::[\w.-]+)?="[^"]*"/g, '')
              .replace('cx="5486400"', `cx="${scenario === 'resize' ? 2743200 : 5486400}"`)
          );
          expect(smartArtNamespaces(xml)).toEqual(['http://schemas.openxmlformats.org/drawingml/2006/diagram']);
          if (extensionSize) expect(xml).toContain('x'.repeat(extensionSize));
          const relationships = await saved.file('word/_rels/document.xml.rels')!.async('string');
          const targets = smartArtTargets(xml, relationships);
          expect(targets).toHaveLength(4);
          for (const target of targets) expect(saved.file(target)).not.toBeNull();
          for (const [name, original] of source.parts) {
            if (name !== 'word/document.xml') expect(await saved.file(name)!.async('string')).toBe(original);
          }
          const reopened = await openDocx(bytes);
          try {
            const segment = reopened.session.storySegments('body').find(segment => segment.kind === 'embed' && segment.embedKind === 'shape');
            if (segment?.kind !== 'embed') throw new Error('saved SmartArt shape missing');
            expect(JSON.parse(segment.payload.shapeJson as string).size.width).toBe(scenario === 'resize' ? 2743200 : 5486400);
            if (scenario !== 'resize') expect(reopened.grep({ query: 'ahead of' }).matches).toHaveLength(1);
          } finally { reopened.close(); }
        } finally { doc.close(); }
      }, 30_000);
    }
  }
}

test('unsupported SmartArt geometry edits warn while saving supported frame changes', async () => {
  const source = smartArtFixture(false);
  const zip = new JSZip();
  for (const [name, xml] of source.parts) zip.file(name, xml);
  const doc = await openDocx(await zip.generateAsync({ type: 'uint8array' }));
  const warning = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    const segment = doc.session.storySegments('body').find(segment => segment.kind === 'embed' && segment.embedKind === 'shape');
    if (segment?.kind !== 'embed') throw new Error('expected SmartArt');
    const shape = JSON.parse(segment.payload.shapeJson as string);
    shape.shapeType = 'ellipse';
    shape.size.width = 2743200;
    doc.session.applyRawOps('body', [{ op: 'setEmbedAttr', index: 24, key: 'shapeJson', value: JSON.stringify(shape) }]);
    const saved = await JSZip.loadAsync(await doc.export());
    const xml = await saved.file('word/document.xml')!.async('string');
    expect(xml).toContain('cx="2743200"');
    expect(smartArtTargets(xml, await saved.file('word/_rels/document.xml.rels')!.async('string'))).toHaveLength(4);
    expect(xml).not.toContain('<wps:wsp');
    expect(warning).toHaveBeenCalledWith('SmartArt diagram content or geometry changes cannot be represented; preserving source diagram content and saving supported frame changes.');
  } finally {
    warning.mockRestore();
    doc.close();
  }
}, 30_000);
