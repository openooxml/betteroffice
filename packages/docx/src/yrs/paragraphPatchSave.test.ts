import { expect, it, spyOn } from 'bun:test';

import type { LegacyParagraphV1, ParagraphPatchV1 } from '../docx/paragraphPatch';
import * as rustWriter from '../docx/rustSaveFacade';
import type { Document } from '../types/document';
import type { YrsSession } from './index';
import type { DocxParagraphSavePlan } from './paragraphIdentity';
import { captureSessionSave, writeSessionSave } from './saveYrsDocx';
import { saveEditorDocument } from './editorSave';

for (const full of [false, true]) {
  it(`accounts for paragraph patches on ${full ? 'full editor' : 'session'} saves`, async () => {
    const paragraphPatches: ParagraphPatchV1[] = [{
      ordinal: 0,
      sessionKey: 'p1',
      seed: { units: [{ kind: 'text', text: 'abc', attrs: {} }], properties: {}, unitSha256: 'seed' },
      textEdits: [{
        from: 1,
        to: 2,
        insert: [{ text: 'X', attrs: {}, inherit: {} }],
        donor: { kind: 'at', at: 1 },
      }],
      runEdits: [],
      paraDeltas: [],
      paraInherit: {},
      resultSha256: 'result',
    }];
    const legacyParagraphs: LegacyParagraphV1[] = [{
      ordinal: 1,
      code: 'UNSUPPORTED_KEY',
      detail: 'key:unknown mark',
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
