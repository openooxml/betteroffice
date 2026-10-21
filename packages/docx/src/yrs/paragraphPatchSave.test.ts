import { expect, it, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';

import type { LegacyParagraphV1, ParagraphPatchV1 } from '../docx/paragraphPatch';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import * as rustWriter from '../docx/rustSaveFacade';
import { preloadOpcWasm, unzipContainer } from '../docx/wasm';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import type { DocxParagraphSavePlan } from './paragraphIdentity';
import { captureSessionSave, saveYrsDocx, writeSessionSave } from './saveYrsDocx';
import { saveEditorDocument } from './editorSave';

for (const full of [false, true]) {
  const behavior = full
    ? "logs unspliced paragraphs under today's full editor save policy"
    : 'forwards paragraph patches on session saves';
  it(behavior, async () => {
    const paragraphPatches: ParagraphPatchV1[] = [{
      ordinal: 0,
      sessionKey: 'p1',
      seed: {
        units: [{ kind: 'text', text: 'abc', attrs: {} }], properties: {}, unitSha256: 'seed',
        inherit: { paragraph: {}, runs: [{}], paragraphMark: {} },
      },
      textEdits: [{
        from: 1,
        to: 2,
        insert: [{ text: 'X', attrs: {}, inherit: {} }],
        donor: { kind: 'at', at: 1 },
      }],
      runEdits: [],
      paraDeltas: [],
      paraInherit: {},
      markInherit: {},
      resultSha256: 'result',
    }];
    const legacyParagraphs: LegacyParagraphV1[] = [{
      ordinal: 1,
      code: 'UNSUPPORTED_STRUCTURE',
      detail: 'restyle: paragraph inheritance',
    }];
    const plan: DocxParagraphSavePlan = {
      assignments: [],
      patchedParts: [],
      splicedParts: [{
        part: 'word/header1.xml',
        sha256: 'part',
        paragraphs: [[0, 'p1'], [1, 'p2']],
        changed: [0, 1],
        paragraphPatches,
        legacyParagraphs,
      }],
    };
    const session = {
      paragraphSavePlan: () => plan,
      paragraphIdentities: () => ({ sessionId: 'session', paragraphs: [] }),
      writtenParagraphIds: () => ({}),
      recordSavedParagraphIds: () => [],
    } as unknown as YrsSession;
    const document = { package: { document: { content: [] } } } as unknown as Document;
    const buffer = new ArrayBuffer(0);
    const writer = spyOn(rustWriter, 'writeDocumentWithRust').mockResolvedValue({
      buffer,
      determinism: { seed: 'seed', now: '2026-10-21T00:00:00Z' },
      warnings: [],
    });
    const diagnostic = spyOn(console, 'debug').mockImplementation(() => {});
    try {
      const capture = captureSessionSave(session);
      if (full) {
        await saveEditorDocument(session, { ...document, originalBuffer: buffer }, []);
        expect(writer.mock.calls[0]![5]!.patchedParts).toEqual([]);
        expect(writer.mock.calls[0]![5]!.splicedParts).toBeUndefined();
        expect(diagnostic).toHaveBeenCalledWith(
          'DOCX_MINIMAL_PATCH unspliced paragraphs',
          [{ part: 'word/header1.xml', ordinals: [0, 1] }]
        );
      } else {
        await writeSessionSave(session, document, capture, buffer, {}, () => false, true, () => true);
        const forwarded = writer.mock.calls[0]![5]!.splicedParts![0]!;
        expect(forwarded.paragraphs).toEqual([0, 1]);
        expect(forwarded.paragraphPatches).toBe(paragraphPatches);
        expect(forwarded.legacyParagraphs).toBe(legacyParagraphs);
        expect(diagnostic).not.toHaveBeenCalled();
      }
    } finally {
      diagnostic.mockRestore();
      writer.mockRestore();
    }
  });
}

it('logs replacement counts per source-opened Yrs save alongside a patched neighbour', async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(
    new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url)
  )));
  await preloadOpcWasm();
  const office = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const source = new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
    )],
    ['_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="doc" Type="${office}/officeDocument" Target="word/document.xml"/>` +
      '</Relationships>'
    )],
    ['word/document.xml', toBytes(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
      '<w:p><w:r><w:t>Field </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>' +
      '<w:p><w:r><w:t>Neighbour</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
    )],
  ])));
  const session = await createYrsSession();
  const diagnostic = spyOn(console, 'debug').mockImplementation(() => {});
  try {
    session.seedFromDocx(source);
    const [field, neighbour] = session.paragraphs('body');
    for (const paragraph of [field!, neighbour!]) {
      session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Edited ');
    }
    const part = session.paragraphSavePlan().splicedParts!
      .find(({ part }) => part === 'word/document.xml')!;
    expect(part.paragraphPatches?.map(({ ordinal }) => ordinal)).toEqual([1]);
    expect(part.legacyParagraphs).toEqual([expect.objectContaining({
      ordinal: 0,
      detail: expect.stringContaining('field:'),
    })]);
    for (let save = 0; save < 2; save += 1) {
      diagnostic.mockClear();
      const saved = await saveYrsDocx(session);
      const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
      expect(xml).toContain('Edited Neighbour');
      expect(diagnostic.mock.calls.filter(([warning]) =>
        typeof warning === 'string' && warning.startsWith('DOCX_MINIMAL_PATCH replace counts')
      )).toEqual([['DOCX_MINIMAL_PATCH replace counts {"field":1}']]);
    }
  } finally {
    diagnostic.mockRestore();
    session.destroy();
  }
});
