import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { openDocx } from '../src';
import { saveEditorDocument, mergeDocxHostMetadata, yrsToDocument } from '@betteroffice/docx/yrs';
import { smartArtDrawingRun, smartArtFixture, smartArtTargets } from '../../docx/src/yrs/__fixtures__/smartArt';

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
