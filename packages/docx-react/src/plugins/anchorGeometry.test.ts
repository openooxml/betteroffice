import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import {
  computeAnchorDisplayTarget,
  createYrsSession,
  proposalSetIdentity,
  type DocxProposalRecord,
  type DocxTextRange,
  type ResidentDocumentRead,
  type ResidentEngineWorkerClient,
  type ResidentProposalReply,
  type YrsSession,
  type YrsStorySegment,
} from '@betteroffice/docx/yrs';
import {
  createResidentEngineSession,
  type ResidentEngineSession,
} from '@betteroffice/docx/yrs/residentEngineSession';
import type { WorkerOpenedDocument } from '../components/DocxEditor/hooks/useDisplayList';
import { revisionPreviewKey } from '../components/DocxEditor/internals/layoutProvenance';
import { holdWorkerOpenDocument } from '../components/DocxEditor/internals/workerOpenReplica';
import { registerWorkerProposalAuthority } from '../components/DocxEditor/internals/workerProposalAuthority';
import {
  hiddenRanges,
  readWorkerAnchorTarget,
  resolveAnchorTarget,
  textRangeToRaw,
} from './anchorGeometry';
import { currentPreviewKey, proposalSnapshot, revisionPreviewOf } from './proposalPreview';

function range(
  start: number,
  end: number,
  view: DocxTextRange['view'] = 'accepted',
  paraId = 'p',
  story = 'body'
): DocxTextRange {
  return { story, start: { paraId, offset: start }, end: { paraId, offset: end }, view };
}

function offsets(segments: readonly YrsStorySegment[], target: DocxTextRange) {
  const result = textRangeToRaw(segments, target);
  if (!result.ok) throw new Error(result.failure.message);
  return [result.range.start.offset, result.range.end.offset];
}

const mark = (paraId: string): YrsStorySegment => ({
  kind: 'pilcrow',
  paraId,
  attributes: {},
  properties: {},
});
const text = (value: string, attributes = {}): YrsStorySegment => ({
  kind: 'text',
  text: value,
  attributes,
});
const embed = (embedKind: string, attributes = {}): YrsStorySegment => ({
  kind: 'embed',
  embedKind,
  attributes,
  payload: {},
});

describe('view-to-raw boundaries', () => {
  test('skips hidden starts, retains ends before hidden text and collapses to one point', () => {
    const segments = [
      text('x', { del: {} }),
      text('ab'),
      text('yy', { del: {} }),
      text('cd'),
      text('z', { del: {} }),
      mark('p'),
    ];
    expect(offsets(segments, range(0, 2))).toEqual([1, 3]);
    expect(offsets(segments, range(2, 4))).toEqual([5, 7]);
    expect(offsets(segments, range(2, 2))).toEqual([5, 5]);
    expect(offsets(segments, range(4, 4))).toEqual([8, 8]);
    expect(offsets([text('gone', { del: {} }), mark('p')], range(0, 0))).toEqual([4, 4]);
  });

  test('counts leading block embeds in offsets but not in the text, and later embeds as atoms', () => {
    for (const kind of ['table', 'blockSdt', 'pageBreak', 'columnBreak']) {
      const segments = [embed(kind), embed(kind), text('😀'), embed('image'), embed(kind), mark('p')];
      expect(offsets(segments, range(0, 4))).toEqual([2, 6]);
      expect(offsets(segments, range(2, 3))).toEqual([4, 5]);
      expect(offsets(segments, range(0, 0))).toEqual([2, 2]);
      expect(textRangeToRaw(segments, range(0, 5))).toMatchObject({ ok: false });
    }
    const hidden = [embed('table'), text('gone', { del: {} }), text('kept'), mark('p')];
    expect(offsets(hidden, range(0, 4))).toEqual([5, 9]);
    expect(offsets(hidden, range(0, 0, 'original'))).toEqual([1, 1]);
  });

  test('uses non-null revision attributes, including falsy values', () => {
    const segments = [
      text('a', { ins: null, del: undefined }),
      text('b', { ins: false }),
      text('c', { del: 0 }),
      mark('p'),
    ];
    expect(offsets(segments, range(0, 2, 'accepted'))).toEqual([0, 2]);
    expect(offsets(segments, range(1, 2, 'original'))).toEqual([2, 3]);
  });

  test('maps cross-paragraph ranges and refuses missing, reversed or invalid bounds', () => {
    const segments = [text('ab'), mark('p'), text('cd'), mark('q')];
    const target = { ...range(1, 1), end: { paraId: 'q', offset: 1 } };
    expect(textRangeToRaw(segments, target)).toEqual({
      ok: true,
      range: {
        start: { story: 'body', paraId: 'p', offset: 1 },
        end: { story: 'body', paraId: 'q', offset: 1 },
      },
    });
    for (const invalid of [
      range(-1, 0),
      range(0, 3),
      range(0.5, 1),
      range(2, 1),
      range(0, 0, 'accepted', 'absent'),
      { ...target, start: target.end, end: target.start },
    ]) {
      expect(textRangeToRaw(segments, invalid)).toMatchObject({
        ok: false,
        failure: { code: 'missing-target' },
      });
    }
    expect(textRangeToRaw([mark('p'), mark('p')], range(0, 0))).toMatchObject({
      ok: false,
      failure: { code: 'ambiguous-target' },
    });
  });
});

async function trackedDocument() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  );
  const runs =
    '<w:r><w:t>A</w:t></w:r><w:del w:id="1" w:author="Reviewer"><w:r><w:delText>dd</w:delText></w:r></w:del><w:r><w:t>B</w:t></w:r><w:ins w:id="2" w:author="Reviewer"><w:r><w:t>ii</w:t></w:r></w:ins><w:r><w:t>C</w:t><w:br/><w:t>D😀</w:t></w:r>';
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001">${runs}</w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr/><w:p w14:paraId="00000002">${runs}</w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="00000003"><w:r><w:t>AB</w:t></w:r><w:ins w:id="3" w:author="Reviewer"><w:r><w:t>xy</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>`
  );
  return zip.generateAsync({ type: 'uint8array' });
}

describe('view-to-raw boundaries in a real session', () => {
  beforeAll(async () => {
    const { preloadEditWasm } = await import('@betteroffice/docx/wasm/edit');
    await preloadEditWasm(
      new Uint8Array(
        readFileSync(
          resolve(import.meta.dir, '../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
        )
      )
    );
  });

  test('maps accepted and original text across revisions and an atom in body and table cells', async () => {
    const session = await createYrsSession({ clientId: 813 });
    try {
      session.openDocx(await trackedDocument(), true);
      for (const [story, paraId] of [
        ['body', '00000001'],
        ['body:t0:r0c0', '00000002'],
      ]) {
        const segments = session.storySegments(story);
        for (const view of ['accepted', 'original'] as const) {
          const read = session.readParagraphs({ story, paraIds: [paraId], view });
          expect(read).toMatchObject({
            ok: true,
            paragraphs: [{ text: view === 'accepted' ? 'ABiiC\uFFFCD😀' : 'AddBC\uFFFCD😀' }],
          });
          expect(offsets(segments, range(0, 9, view, paraId, story))).toEqual([0, 11]);
          expect(offsets(segments, range(4, 5, view, paraId, story))).toEqual([6, 7]);
          expect(offsets(segments, range(5, 6, view, paraId, story))).toEqual([7, 8]);
          expect(offsets(segments, range(6, 9, view, paraId, story))).toEqual([8, 11]);
          expect(textRangeToRaw(segments, range(0, 10, view, paraId, story))).toMatchObject({
            ok: false,
            failure: { code: 'missing-target' },
          });
        }
        expect(offsets(segments, range(0, 1, 'accepted', paraId, story))).toEqual([0, 1]);
        expect(offsets(segments, range(1, 1, 'accepted', paraId, story))).toEqual([3, 3]);
        expect(offsets(segments, range(1, 2, 'accepted', paraId, story))).toEqual([3, 4]);
        expect(offsets(segments, range(3, 4, 'original', paraId, story))).toEqual([3, 4]);
        expect(offsets(segments, range(4, 4, 'original', paraId, story))).toEqual([6, 6]);
      }
      const segments = session.storySegments('body');
      const found = session.findText({
        text: 'xy',
        within: { kind: 'paragraph', story: 'body', paraId: '00000003' },
        view: 'accepted',
      });
      if (!found.ok) throw new Error(found.failure.message);
      const inserted = session
        .listRevisions()
        .find((revision) => revision.range.start.paraId === '00000003')!;
      expect(textRangeToRaw(segments, found.matches[0]!.range)).toEqual({
        ok: true,
        range: {
          start: { story: 'body', ...inserted.range.start },
          end: { story: 'body', ...inserted.range.end },
        },
      });
      expect(textRangeToRaw(segments, range(0, 0, 'accepted', 'missing'))).toMatchObject({
        ok: false,
        failure: { code: 'missing-target' },
      });
    } finally {
      session.destroy();
    }
  });

  test('reads each whole story once per document version', async () => {
    const session = await createYrsSession({ clientId: 814 });
    try {
      session.openDocx(await trackedDocument(), true);
      const reads = new Map<string, number>();
      const counted = new Proxy(session, {
        get(target, key, receiver) {
          const value = Reflect.get(target, key, receiver);
          if (typeof value !== 'function') return value;
          return (...args: unknown[]) => {
            reads.set(String(key), (reads.get(String(key)) ?? 0) + 1);
            return value.apply(target, args);
          };
        },
      });
      const paragraph = {
        kind: 'persisted' as const,
        story: { partUri: '/word/document.xml', kind: 'body' as const },
        paraId: '00000003',
      };
      const revisionId = session.listRevisions()[0]!.revisionId;
      const resolveAll = () =>
        [
          { kind: 'revision' as const, revisionId },
          { kind: 'paragraph' as const, paragraph },
          { kind: 'search' as const, paragraph, text: 'xy' },
        ].map((target) => resolveAnchorTarget(counted, target, session.version()));
      const first = resolveAll();
      expect(first.every((resolved) => resolved.ok)).toBe(true);
      expect(resolveAll()).toEqual(first);
      expect(reads.get('listRevisions')).toBe(1);
      expect(reads.get('paragraphSpans')).toBe(1);
      expect(reads.get('storySegments')).toBe(1);
      expect(reads.get('resolveParagraphAnchor')).toBe(1);

      session.insertText({ story: 'body', paraId: '00000003', offset: 0 }, 'z');
      const changed = resolveAll();
      expect(reads.get('listRevisions')).toBe(2);
      expect(reads.get('storySegments')).toBe(2);
      const offset = (resolved: (typeof first)[number]) =>
        resolved.ok ? resolved.ranges[0]!.start.offset : null;
      expect(offset(changed[2]!)).toBe(offset(first[2]!)! + 1);
    } finally {
      session.destroy();
    }
  });
});

function residentRead(engine: ResidentEngineSession) {
  const requests: ResidentDocumentRead[] = [];
  const read = (async (request: ResidentDocumentRead) => {
    requests.push(request);
    const version = engine.proposalEngine.version();
    if (request.kind === 'resolveParagraphAnchors') {
      return {
        version,
        value: { results: request.anchors.map((anchor) => engine.geometryReader.resolveParagraphAnchor(anchor)) },
      };
    }
    if (request.kind !== 'anchorTarget') throw new Error(`unexpected ${request.kind} read`);
    return {
      version,
      value:
        version === request.expectVersion
          ? computeAnchorDisplayTarget(engine.geometryReader, request.target, request.revisionPreview)
          : null,
    };
  }) as ResidentEngineWorkerClient['documentRead'];
  return { read, requests };
}

describe('anchor reads through the document worker', () => {
  beforeAll(async () => {
    const { preloadEditWasm } = await import('@betteroffice/docx/wasm/edit');
    await preloadEditWasm(
      new Uint8Array(
        readFileSync(
          resolve(import.meta.dir, '../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
        )
      )
    );
  });

  const persisted = {
    kind: 'persisted' as const,
    story: { partUri: '/word/document.xml', kind: 'body' as const },
    paraId: '00000003',
  };

  test('a worker bootstrapped from main state at another version answers every target kind like the main thread, persisted anchors resolved on main', async () => {
    const main = await createYrsSession({ clientId: 815 });
    const engine = await createResidentEngineSession(undefined, 816);
    try {
      main.openDocx(await trackedDocument(), true);
      engine.loadState(main.encodeState());
      expect(engine.geometryReader.resolveParagraphAnchor(persisted)).toMatchObject({
        status: 'unsupported',
      });
      const { read, requests } = residentRead(engine);
      const found = main.findText({
        text: 'xy',
        within: { kind: 'paragraph', story: 'body', paraId: '00000003' },
        view: 'accepted',
      });
      if (!found.ok) throw new Error(found.failure.message);
      const resolved = main.resolveParagraphAnchor({ ...persisted, paraId: '00000001' });
      if (resolved.status !== 'found' || resolved.anchor.kind !== 'session') throw new Error('expected a session anchor');
      const session = resolved.anchor;
      const range = { kind: 'range' as const, version: main.version(), range: found.matches[0]!.range };
      const key = currentPreviewKey(main);
      const targets = [
        { kind: 'paragraph' as const, paragraph: persisted },
        { kind: 'paragraph' as const, paragraph: session },
        { kind: 'search' as const, paragraph: persisted, text: 'xy' },
        { kind: 'revision' as const, revisionId: main.listRevisions()[0]!.revisionId },
        range,
      ];
      for (const target of targets) {
        const sent = target.kind === 'range' ? { ...target, version: engine.proposalEngine.version() } : target;
        const expected = computeAnchorDisplayTarget(main, target, undefined);
        expect(expected).toMatchObject({ ok: true, ranges: expect.arrayContaining([expect.any(Object)]) });
        expect(await readWorkerAnchorTarget(read, main, sent, engine.proposalEngine.version(), 0, key)).toEqual(expected);
      }
      expect(engine.proposalEngine.version()).not.toBe(main.version());
      expect(requests.map((request) => request.kind)).toEqual(targets.map(() => 'anchorTarget'));
      expect(await readWorkerAnchorTarget(read, main, range, engine.proposalEngine.version(), 0, key)).toMatchObject({
        ok: false,
        failure: { code: 'stale-version' },
      });
      expect(requests).toContainEqual(
        expect.objectContaining({ target: expect.objectContaining({ paragraph: expect.objectContaining({ kind: 'session', paraId: '00000003' }) }) })
      );
      expect(await readWorkerAnchorTarget(read, main, targets[3]!, 'superseded', 0, key)).toBeNull();
      expect(await readWorkerAnchorTarget(read, main, targets[3]!, engine.proposalEngine.version(), 1, key)).toMatchObject({
        ok: false,
        failure: { code: 'layout-unavailable' },
      });
    } finally {
      main.destroy();
      engine.destroy();
    }
  });

  test('a viewer whose document stays in the worker resolves persisted anchors there', async () => {
    const main = await createYrsSession({ clientId: 817 });
    const engine = await createResidentEngineSession(undefined, 818);
    const reference = await createYrsSession({ clientId: 819 });
    try {
      const bytes = await trackedDocument();
      engine.openDocx(bytes);
      reference.openDocx(bytes, true);
      holdWorkerOpenDocument(main, () => {
        throw new Error('released');
      });
      const { read, requests } = residentRead(engine);
      const target = { kind: 'search' as const, paragraph: persisted, text: 'xy' };
      const expected = computeAnchorDisplayTarget(reference, target, undefined);
      expect(expected).toMatchObject({ ok: true });
      expect(await readWorkerAnchorTarget(read, main, target, engine.proposalEngine.version(), 0, currentPreviewKey(main))).toEqual(expected);
      expect(requests.map((request) => request.kind)).toEqual(['resolveParagraphAnchors', 'anchorTarget']);
      expect(await readWorkerAnchorTarget(read, main, target, 'superseded', 0, currentPreviewKey(main))).toBeNull();
      expect(requests.at(-1)?.kind).toBe('resolveParagraphAnchors');
      expect(
        await readWorkerAnchorTarget(read, main, { ...target, paragraph: { ...persisted, paraId: '0000FFFF' } }, engine.proposalEngine.version(), 0, currentPreviewKey(main))
      ).toMatchObject({ ok: false, failure: { code: 'missing-target' } });
    } finally {
      main.destroy();
      engine.destroy();
      reference.destroy();
    }
  });
});

describe('revision preview of an editor peer', () => {
  const record = (id: string, state: DocxProposalRecord['state'], revisionId: string): DocxProposalRecord => ({
    id,
    state,
    paragraph: { kind: 'session', sessionId: 's', story: 'body', paraId: 'p' },
    revisionIds: [revisionId],
    changed: true,
  });

  test('hidden ranges, the preview key and worker anchor reads share the merged preview', async () => {
    let local = { version: 'worker-1', previewVersion: 1, proposals: [record('local', 'rejected', 'r2')] };
    const revision = (revisionId: string, kind: 'insertion' | 'deletion', offset: number) => ({
      revisionId,
      kind,
      story: 'body',
      range: { start: { paraId: 'p', offset }, end: { paraId: 'p', offset: offset + 1 } },
    });
    const session = {
      version: () => 'worker-1',
      encodeStateVector: () => new Uint8Array(),
      getProposals: () => local,
      listRevisions: () => [revision('r1', 'deletion', 0), revision('r2', 'insertion', 2), revision('r3', 'deletion', 4)],
      storiesChangedSince: () => ({ revision: 0, stories: [] }),
    } as unknown as YrsSession;
    const worker = record('worker', 'accepted', 'r1');
    const snapshot = { version: 'worker-1', previewVersion: 1, proposals: [worker] };
    const reply: ResidentProposalReply = {
      mirror: {
        version: 'worker-1',
        proposals: { previewVersion: 1, entries: [{ record: worker, key: 'k', suggest: { author: 'a', date: 'd' } }] },
      },
      result: { ok: true, snapshot },
      changedStories: [],
      geometry: { version: 'worker-1', previewVersion: 1, proposals: proposalSetIdentity(snapshot), targets: {}, hidden: [] },
      updates: [],
      stateVector: new Uint8Array(),
    };
    const requests: ResidentDocumentRead[] = [];
    const authority = registerWorkerProposalAuthority(
      session,
      {
        proposal: async () => reply,
        documentRead: async () => {
          throw new Error('unexpected authority read');
        },
        handOver: async () => {
          throw new Error('unexpected handover');
        },
      } as unknown as WorkerOpenedDocument,
      {
        editorPeer: true,
        current: () => true,
        laidOut: async () => {},
        relayout: () => {},
        adopted: () => {},
        contentChanged: () => {},
      }
    );
    await authority.initialize();
    const merged = { r1: 'accepted', r2: 'rejected' } as const;
    expect(revisionPreviewOf(session)).toEqual(merged);
    expect(currentPreviewKey(session)).toBe(revisionPreviewKey(merged));
    expect(hiddenRanges(session, 'worker-1')).toEqual([
      { start: { story: 'body', paraId: 'p', offset: 0 }, end: { story: 'body', paraId: 'p', offset: 1 } },
      { start: { story: 'body', paraId: 'p', offset: 2 }, end: { story: 'body', paraId: 'p', offset: 3 } },
    ]);
    const read = (async (request: ResidentDocumentRead) => {
      requests.push(request);
      return { version: 'worker-1', value: { ok: true, ranges: [], paragraph: 0, hidden: [] } };
    }) as ResidentEngineWorkerClient['documentRead'];
    const previewVersion = proposalSnapshot(session)!.previewVersion;
    const key = currentPreviewKey(session);
    expect(
      await readWorkerAnchorTarget(read, session, { kind: 'revision', revisionId: 'r3' }, 'worker-1', previewVersion, key)
    ).toMatchObject({ ok: true });
    expect(requests).toEqual([
      { kind: 'anchorTarget', target: { kind: 'revision', revisionId: 'r3' }, revisionPreview: merged, expectVersion: 'worker-1' },
    ]);
    local = { ...local, proposals: [record('local', 'accepted', 'r2')] };
    expect(proposalSnapshot(session)!.previewVersion).toBe(previewVersion);
    expect(currentPreviewKey(session)).not.toBe(key);
    expect(
      await readWorkerAnchorTarget(read, session, { kind: 'revision', revisionId: 'r3' }, 'worker-1', previewVersion, key)
    ).toMatchObject({ ok: false, failure: { code: 'layout-unavailable' } });
    expect(requests).toHaveLength(1);
  });
});
