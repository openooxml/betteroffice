import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import type { DocxProposalInput, DocxProposalResult } from './proposals';
import {
  computeProposalGeometryMirror,
  proposalSetIdentity,
  resolveNavigationTarget,
  resolveMirroredNavigationTarget,
  type ProposalGeometryReader,
} from './proposalGeometry';
import { createResidentEngineSession } from './residentEngineSession';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const BODY = { partUri: '/word/document.xml', kind: 'body' } as const;
const SUGGEST = { author: 'Reviewer', date: '2026-09-29T12:00:00Z' };

function fixture(): Uint8Array {
  const paragraph = (id: string, text: string) =>
    `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const body = [
    paragraph('00000001', 'Keep gone text'),
    paragraph('00000002', 'Insert here'),
    paragraph('00000003', ''),
    paragraph('00000004', 'Join '),
    paragraph('00000005', 'Vanish'),
    '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>',
    paragraph('0000C001', 'cell value'),
    '</w:tc></w:tr></w:tbl>',
    paragraph('00000006', 'Tail'),
    paragraph('00000007', 'Unchanged'),
  ].join('');
  const parts: PartsMap = new Map();
  parts.set(
    '[Content_Types].xml',
    toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
  );
  parts.set(
    '_rels/.rels',
    toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
  );
  parts.set(
    'word/document.xml',
    toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr/></w:body></w:document>`)
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function snapshotOf(result: DocxProposalResult) {
  if (!result.ok) throw new Error(result.failure.message);
  return result.snapshot;
}

function replace(id: string, paraId: string, search: string, replaceWith: string): DocxProposalInput {
  return {
    id,
    paragraph: { kind: 'persisted', story: BODY, paraId },
    suggest: SUGGEST,
    op: 'replaceText',
    search,
    replaceWith,
  };
}

async function proposedDocument(): Promise<YrsSession> {
  const main = await createYrsSession({ clientId: 79101 });
  try {
    main.openDocx(fixture(), true);
    snapshotOf(
      main.proposeChanges({
        expectVersion: main.version(),
        proposals: [
          replace('delete', '00000001', 'gone ', ''),
          replace('insert', '00000002', 'Insert', 'Added'),
          replace('empty', '00000003', '', ''),
          replace('missing', '00000005', 'Vanish', 'Vanish'),
          replace('cell', '0000C001', 'value', 'content'),
          replace('tail', '00000006', 'Tail', 'End'),
          replace('same', '00000007', 'Unchanged', 'Unchanged'),
        ],
      })
    );
    snapshotOf(
      main.setProposalStates({
        expectVersion: main.version(),
        expectPreviewVersion: main.getProposals().previewVersion,
        changes: [
          { id: 'delete', state: 'accepted' },
          { id: 'insert', state: 'rejected' },
        ],
      })
    );
    main.mergeParagraphs('body', '00000004');
    return main;
  } catch (error) {
    main.destroy();
    throw error;
  }
}

function duplicated(reader: ProposalGeometryReader): ProposalGeometryReader {
  return {
    ...reader,
    paragraphs: (story) => {
      const paragraphs = reader.paragraphs(story);
      const paragraph = paragraphs.find(({ paraId }) => paraId === '00000007');
      return paragraph ? [...paragraphs, paragraph] : paragraphs;
    },
    paragraphIdCount: (story, paraId) =>
      reader.paragraphIdCount(story, paraId) + (paraId === '00000007' ? 1 : 0),
    paragraphSpans: (story) => {
      const spans = reader.paragraphSpans(story);
      const span = spans.find(({ paraId }) => paraId === '00000007');
      return span ? [...spans, span] : spans;
    },
  };
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

describe('proposal geometry readers', () => {
  test('can omit navigation targets without reading or building the sidebar projection', async () => {
    const main = await proposedDocument();
    const count = spyOn(main, 'paragraphIdCount');
    const stories = spyOn(main, 'storyIds');
    try {
      const snapshot = main.getProposals();
      const mirror = computeProposalGeometryMirror(main, snapshot, false);
      expect(mirror.navigationTargets).toBeUndefined();
      expect(count).not.toHaveBeenCalled();
      expect(stories).not.toHaveBeenCalled();
      const complete = computeProposalGeometryMirror(main, snapshot);
      expect({ ...complete, navigationTargets: undefined }).toEqual<typeof mirror>(mirror);
      expect(count).toHaveBeenCalled();
      expect(stories).toHaveBeenCalledTimes(1);
    } finally {
      count.mockRestore();
      stories.mockRestore();
      main.destroy();
    }
  });

  test('matches the main session on a resident replica, including hidden and missing targets', async () => {
    const main = await proposedDocument();
    const resident = await createResidentEngineSession();
    try {
      resident.loadState(main.encodeState());
      const snapshot = main.getProposals();
      const expected = computeProposalGeometryMirror(main, snapshot);
      const actual = computeProposalGeometryMirror(resident.geometryReader, snapshot);
      expect(actual.version).toBe(resident.geometryReader.version());
      expect({ ...actual, version: expected.version }).toEqual(expected);
      expect(actual.previewVersion).toBe(snapshot.previewVersion);
      expect(actual.proposals).toBe(proposalSetIdentity(snapshot));
      for (const { id, paragraph } of snapshot.proposals) {
        const navigation = resolveNavigationTarget(main, paragraph.story, paragraph.paraId);
        expect(actual.navigationTargets?.[id]).toEqual(navigation);
        expect(resolveMirroredNavigationTarget(
          { ...actual, version: snapshot.version }, snapshot, paragraph.story, paragraph.paraId
        )).toEqual(navigation);
      }
      expect(Object.keys(actual.targets)).toEqual(snapshot.proposals.map(({ id }) => id));
      expect(actual.hidden).toHaveLength(2);
      expect(actual.hidden.every(({ from, to }) => from < to)).toBe(true);
      expect(actual.targets.missing).toEqual({
        ok: false,
        failure: { code: 'missing-target', message: 'The paragraph no longer exists' },
      });
      for (const id of ['empty', 'same']) {
        expect(actual.targets[id]).toMatchObject({
          ok: true,
          ranges: [],
          paragraph: expect.any(Number),
        });
      }
      expect(actual.targets.cell).toMatchObject({
        ok: true,
        ranges: [expect.any(Object), expect.any(Object)],
        paragraph: expect.any(Number),
      });
      expect(resident.paragraphIdentities().sessionId).toBe(main.paragraphIdentities().sessionId);
      expect(resident.paragraphIdentities().paragraphs.map(({ session }) => session)).toEqual(
        main.paragraphIdentities().paragraphs.map(({ session }) => session)
      );
      expect(resident.geometryReader.listRevisions()).toEqual(main.listRevisions());

      for (const [story, paraId] of [
        ['body', '00000001'],
        ['body:t0:r0c0', '0000C001'],
        ['body', '00000005'],
        ['absent', '00000001'],
      ]) {
        expect(resolveNavigationTarget(resident.geometryReader, story, paraId)).toEqual(
          resolveNavigationTarget(main, story, paraId)
        );
      }
      expect(resolveNavigationTarget(main, 'body', '00000001')).toMatchObject({
        loc: { story: 'body', paraId: '00000001', offset: 0 },
        position: 1,
      });
      expect(resolveNavigationTarget(main, 'body', '00000005')).toBe('missing-target');

      const duplicatedMain = duplicated(main);
      const duplicatedResident = duplicated(resident.geometryReader);
      const duplicateExpected = computeProposalGeometryMirror(duplicatedMain, snapshot);
      const duplicateActual = computeProposalGeometryMirror(duplicatedResident, snapshot);
      expect({ ...duplicateActual, version: duplicateExpected.version }).toEqual(duplicateExpected);
      expect(duplicateActual.targets.same).toEqual({
        ok: false,
        failure: { code: 'ambiguous-target', message: 'The paragraph cannot be resolved uniquely' },
      });
      expect(duplicateActual.navigationTargets?.same).toBe('ambiguous-target');
      expect(resolveNavigationTarget(duplicatedMain, 'body', '00000007')).toBe('ambiguous-target');
      expect(resolveNavigationTarget(duplicatedResident, 'body', '00000007')).toBe('ambiguous-target');
    } finally {
      resident.destroy();
      main.destroy();
    }
  });

  test('uses mirrored navigation only for a matching document, preview and proposal set', async () => {
    const main = await proposedDocument();
    try {
      const snapshot = main.getProposals();
      const mirror = computeProposalGeometryMirror(main, snapshot);
      const paragraph = snapshot.proposals.find(({ id }) => id === 'cell')!.paragraph;
      const resolve = (geometry: typeof mirror | null, current = snapshot) =>
        resolveMirroredNavigationTarget(geometry, current, paragraph.story, paragraph.paraId);
      expect(resolve(mirror)).toEqual(resolveNavigationTarget(main, paragraph.story, paragraph.paraId));
      expect(resolve(null)).toBeNull();
      expect(resolve({ ...mirror, navigationTargets: undefined })).toBeNull();
      expect(resolve({ ...mirror, navigationTargets: {} })).toBeNull();
      expect(resolve({ ...mirror, version: 'older' })).toBeNull();
      expect(resolve({ ...mirror, previewVersion: snapshot.previewVersion + 1 })).toBeNull();
      expect(resolve({ ...mirror, proposals: 'other' })).toBeNull();
      expect(resolve(mirror, { ...snapshot, version: 'newer' })).toBeNull();
      expect(resolveMirroredNavigationTarget(mirror, snapshot, 'body', 'unknown')).toBeNull();
      expect(resolveMirroredNavigationTarget(mirror, snapshot, 'other', paragraph.paraId)).toBeNull();
    } finally {
      main.destroy();
    }
  });

  test('gives equal proposal snapshots the same identity', async () => {
    const main = await proposedDocument();
    try {
      const first = main.getProposals();
      const equal = main.getProposals();
      expect(equal).not.toBe(first);
      expect(equal).toEqual(first);
      expect(proposalSetIdentity(equal)).toBe(proposalSetIdentity(first));
    } finally {
      main.destroy();
    }
  });

  test('changes identity on withdrawal and re-proposal at another paragraph', async () => {
    const main = await createYrsSession({ clientId: 79102 });
    try {
      main.openDocx(fixture(), true);
      const snapshot = snapshotOf(
        main.proposeChanges({
          expectVersion: main.version(),
          proposals: [replace('same', '00000007', 'Unchanged', 'Unchanged')],
        })
      );
      expect(snapshot.proposals.find(({ id }) => id === 'same')).toMatchObject({
        changed: false,
        revisionIds: [],
      });
      const first = computeProposalGeometryMirror(main, snapshot);
      const withdrawn = snapshotOf(
        main.withdrawProposals({ expectVersion: main.version(), ids: ['same'] })
      );
      expect(withdrawn.version).toBe(first.version);
      expect(withdrawn.previewVersion).toBe(first.previewVersion);
      expect(withdrawn.proposals.some(({ id }) => id === 'same')).toBe(false);
      expect(proposalSetIdentity(withdrawn)).not.toBe(first.proposals);

      const proposed = snapshotOf(
        main.proposeChanges({
          expectVersion: main.version(),
          proposals: [replace('same', '00000002', 'Insert here', 'Insert here')],
        })
      );
      const next = computeProposalGeometryMirror(main, proposed);
      expect(next.version).toBe(first.version);
      expect(next.previewVersion).toBe(first.previewVersion);
      expect(next.proposals).toBe(proposalSetIdentity(proposed));
      expect(next.proposals).not.toBe(first.proposals);
      expect(next.proposals).not.toBe(proposalSetIdentity(withdrawn));
      expect(first.targets.same).toEqual({ ok: true, ranges: [], paragraph: 71 });
      expect(next.targets.same).toEqual({ ok: true, ranges: [], paragraph: 17 });
    } finally {
      main.destroy();
    }
  });

  test('recomputes hidden ranges when only the preview decisions change', async () => {
    const main = await proposedDocument();
    try {
      const first = computeProposalGeometryMirror(main, main.getProposals());
      snapshotOf(
        main.setProposalStates({
          expectVersion: main.version(),
          expectPreviewVersion: main.getProposals().previewVersion,
          changes: [
            { id: 'delete', state: 'proposed' },
            { id: 'insert', state: 'proposed' },
          ],
        })
      );
      const next = computeProposalGeometryMirror(main, main.getProposals());
      expect(next.version).toBe(first.version);
      expect(next.previewVersion).toBe(first.previewVersion + 1);
      expect(next.targets).toEqual(first.targets);
      expect(next.hidden).toEqual([]);
    } finally {
      main.destroy();
    }
  });

  test('keeps an unmappable fallback paragraph null for proposals without revisions', async () => {
    const main = await proposedDocument();
    try {
      const reader: ProposalGeometryReader = { ...main, hasStory: () => false };
      const mirror = computeProposalGeometryMirror(reader, main.getProposals());
      expect(mirror.targets.same).toEqual({ ok: true, ranges: [], paragraph: null });
      expect(mirror.targets.empty).toEqual({ ok: true, ranges: [], paragraph: null });
      expect(mirror.hidden).toEqual([]);
    } finally {
      main.destroy();
    }
  });

  test('refuses unmappable target ranges and drops unmappable hidden intervals', async () => {
    const main = await proposedDocument();
    try {
      const reader: ProposalGeometryReader = {
        ...main,
        hasStory: (story) => story !== 'body:t0:r0c0' && main.hasStory(story),
        listRevisions: () =>
          main.listRevisions().map((revision) =>
            revision.kind === 'deletion'
              ? { ...revision, story: 'body:absent' }
              : revision
          ),
      };
      const mirror = computeProposalGeometryMirror(reader, main.getProposals());
      expect(mirror.targets.cell).toEqual({
        ok: false,
        failure: { code: 'unsupported', message: 'The target has no body display position' },
      });
      expect(mirror.targets.delete).toEqual({
        ok: false,
        failure: { code: 'unsupported', message: 'The target has no body display position' },
      });
      expect(mirror.hidden).toHaveLength(1);
    } finally {
      main.destroy();
    }
  });
});
