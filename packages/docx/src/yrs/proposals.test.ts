import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { createProposalRegistry } from './proposals';
import {
  createYrsSession,
  proposalRevisionPreview,
  saveYrsDocx,
  type DocxOccurrence,
  type DocxProposalInput,
  type DocxProposalResult,
  type DocxProposalSnapshot,
  type DocxSourceStory,
  type YrsSession,
} from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE = 'application/vnd.openxmlformats-officedocument';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  `xmlns:r="${REL}"`,
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
].join(' ');
const BODY: DocxSourceStory = { partUri: '/word/document.xml', kind: 'body' };
const SUGGEST = { author: 'Atira', date: '2026-09-29T12:00:00Z' };

const run = (text: string, props = '') =>
  `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (id: string, content: string) => `<w:p w14:paraId="${id}">${content}</w:p>`;
const cell = (content: string) => `<w:tc>${content}</w:tc>`;

const DOCUMENT = [
  paragraph('00000001', `${run('Hello ')}${run('wor', '<w:b/>')}${run('ld')}`),
  paragraph('00000002', run('aaaa and aaaa')),
  paragraph('00000003', run('Keep this sentence.')),
  paragraph('0000000D', run('First twin')),
  paragraph('0000000D', run('Second twin')),
  `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell(
    paragraph('0000C001', run('cell value'))
  )}</w:tr></w:tbl>`,
  paragraph('00000005', run('a '.repeat(200).trimEnd())),
  paragraph('00000006', run('Tail')),
].join('');

const SMALL_MARK = '<w:pPr><w:rPr><w:sz w:val="18"/></w:rPr></w:pPr>';

/** Empty paragraphs to fill, and paragraphs whose only content still counts as text. */
const FILLABLE = [
  `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr>${cell(
    paragraph('0000F001', SMALL_MARK)
  )}</w:tr></w:tbl>`,
  paragraph('0000F002', ''),
  paragraph('0000F003', `<w:hyperlink w:anchor="target">${run('link')}</w:hyperlink>`),
  paragraph('0000F004', `<w:sdt><w:sdtPr/><w:sdtContent>${run('ctl')}</w:sdtContent></w:sdt>`),
  paragraph('0000F005', `<w:fldSimple w:instr=" PAGE ">${run('1')}</w:fldSimple>`),
  paragraph(
    '0000F006',
    `<w:ins w:id="41" w:author="Bob" w:date="2026-09-01T00:00:00Z">${run('new')}</w:ins>`
  ),
  paragraph('0000F007', ''),
  paragraph('0000F007', ''),
  paragraph('0000F008', run('Keep this sentence.')),
  paragraph(
    '0000F009',
    `<w:del w:id="42" w:author="Bob" w:date="2026-09-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>`
  ),
].join('');

function fixture(body = DOCUMENT): Uint8Array {
  const parts: PartsMap = new Map();
  const set = (name: string, content: string) => parts.set(name, toBytes(content));
  set(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.wordprocessingml.document.main+xml"/></Types>`
  );
  set(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`
  );
  set('word/document.xml', `<w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const sessions: YrsSession[] = [];
let nextClientId = 78100;

async function open(bytes = fixture()): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: nextClientId++ });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

const at = (paraId: string) => ({ kind: 'persisted', story: BODY, paraId } as const);

function replace(
  id: string,
  paraId: string,
  search: string,
  replaceWith: string,
  occurrence?: DocxOccurrence
): DocxProposalInput {
  return {
    id,
    paragraph: at(paraId),
    suggest: SUGGEST,
    op: 'replaceText',
    search,
    replaceWith,
    ...(occurrence === undefined ? {} : { occurrence }),
  };
}

function insert(
  id: string,
  paraId: string,
  position: 'start' | 'end' | { offset: number },
  text: string
): DocxProposalInput {
  return { id, paragraph: at(paraId), suggest: SUGGEST, op: 'insertText', at: position, text };
}

function snapshotOf(result: DocxProposalResult): DocxProposalSnapshot {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result.snapshot;
}

function texts(session: YrsSession, view: 'accepted' | 'original', story = 'body'): string[] {
  const read = session.readParagraphs({ story, view });
  if (!read.ok) throw new Error(read.failure.message);
  return read.paragraphs.map((entry) => entry.text);
}

function propose(session: YrsSession, ...proposals: DocxProposalInput[]): DocxProposalResult {
  return session.proposeChanges({ expectVersion: session.version(), proposals });
}

function decide(
  session: YrsSession,
  changes: Array<{ id: string; state: 'proposed' | 'accepted' | 'rejected' }>
): DocxProposalResult {
  return session.setProposalStates({
    expectVersion: session.version(),
    expectPreviewVersion: session.getProposals().previewVersion,
    changes,
  });
}

/** The document, history and registry, to show a refusal changed none of them. */
function state(session: YrsSession) {
  return {
    version: session.version(),
    revisions: session.listRevisions(),
    canUndo: session.canUndo(),
    canRedo: session.canRedo(),
    proposals: session.getProposals(),
  };
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

describe('YrsSession host proposals', () => {
  it('proposes a cross-run replacement as tracked revisions it records', async () => {
    const session = await open();
    const before = session.version();
    const snapshot = snapshotOf(propose(session, replace('p1', '00000001', 'world', 'earth')));
    expect(snapshot.version).not.toBe(before);
    expect(snapshot.version).toBe(session.version());
    expect(snapshot.previewVersion).toBe(0);
    const [record] = snapshot.proposals;
    expect(record).toMatchObject({
      id: 'p1',
      state: 'proposed',
      changed: true,
      paragraph: { kind: 'session', story: 'body', paraId: '00000001' },
    });
    expect(record!.revisionIds.length).toBeGreaterThan(0);
    const listed = new Set(session.listRevisions().map((revision) => revision.revisionId));
    for (const id of record!.revisionIds) expect(listed.has(id)).toBe(true);
    expect(texts(session, 'accepted')[0]).toBe('Hello earth');
    expect(texts(session, 'original')[0]).toBe('Hello world');
    const revised = session
      .storySegments('body')
      .filter(
        (segment) =>
          segment.kind === 'text' && (segment.attributes.del || segment.attributes.ins)
      );
    expect(revised.map((segment) => (segment.kind === 'text' ? segment.text : '')).join('')).toBe(
      'worldearth'
    );
    expect(revised[0]!.attributes.del).toMatchObject({ author: 'Atira' });
    expect(revised.at(-1)!.attributes.ins).toMatchObject({ author: 'Atira' });
    expect(session.listRevisions().every((revision) => revision.author === 'Atira')).toBe(true);
  });

  it('proposes pure deletions and insertions at start, end and an offset', async () => {
    const session = await open();
    const snapshot = snapshotOf(
      propose(
        session,
        replace('delete', '00000003', ' this', ''),
        insert('start', '00000006', 'start', '['),
        insert('end', '00000006', 'end', ']'),
        insert('offset', '00000001', { offset: 5 }, ',')
      )
    );
    expect(snapshot.proposals.map((proposal) => [proposal.id, proposal.changed])).toEqual([
      ['delete', true],
      ['start', true],
      ['end', true],
      ['offset', true],
    ]);
    const accepted = texts(session, 'accepted');
    expect(accepted[0]).toBe('Hello, world');
    expect(accepted[2]).toBe('Keep sentence.');
    expect(accepted.at(-1)).toBe('[Tail]');
    expect(texts(session, 'original')[2]).toBe('Keep this sentence.');
    const kinds = new Map(
      session.listRevisions().map((revision) => [revision.revisionId, revision.kind])
    );
    const kindsOf = (id: string) =>
      snapshot.proposals
        .find((proposal) => proposal.id === id)!
        .revisionIds.map((revision) => kinds.get(revision));
    expect(kindsOf('delete')).toEqual(['deletion']);
    expect(kindsOf('start')).toEqual(['insertion']);
    expect(propose(session, insert('far', '00000006', { offset: 99 }, '!'))).toMatchObject({
      ok: false,
      failure: { code: 'invalid-step', proposalId: 'far' },
    });
    expect(propose(session, insert('negative', '00000006', { offset: -1 }, '!'))).toMatchObject({
      ok: false,
      failure: { code: 'invalid-step', proposalId: 'negative' },
    });
  });

  it('selects the first, nth or every non-overlapping occurrence', async () => {
    const first = await open();
    snapshotOf(propose(first, replace('first', '00000002', 'aa', 'b')));
    expect(texts(first, 'accepted')[1]).toBe('baa and aaaa');

    const second = await open();
    snapshotOf(propose(second, replace('second', '00000002', 'aa', 'b', 2)));
    expect(texts(second, 'accepted')[1]).toBe('aab and aaaa');

    const fourth = await open();
    snapshotOf(propose(fourth, replace('fourth', '00000002', 'aa', 'b', 4)));
    expect(texts(fourth, 'accepted')[1]).toBe('aaaa and aab');

    const all = await open();
    const snapshot = snapshotOf(propose(all, replace('all', '00000002', 'aa', 'b', 'all')));
    expect(texts(all, 'accepted')[1]).toBe('bb and bb');
    expect(snapshot.proposals[0]!.revisionIds).toHaveLength(2);

    const missing = await open();
    const unchanged = state(missing);
    expect(propose(missing, replace('fifth', '00000002', 'aa', 'b', 5))).toMatchObject({
      ok: false,
      failure: { code: 'missing-target', proposalId: 'fifth' },
    });
    expect(propose(missing, replace('absent', '00000002', 'zz', 'b'))).toMatchObject({
      ok: false,
      failure: { code: 'missing-target', proposalId: 'absent' },
    });
    for (const occurrence of [0, 1.5, -1]) {
      expect(propose(missing, replace('bad', '00000002', 'aa', 'b', occurrence))).toMatchObject({
        ok: false,
        failure: { code: 'invalid-step', proposalId: 'bad' },
      });
    }
    expect(propose(missing, replace('empty', '00000002', '', 'b'))).toMatchObject({
      ok: false,
      failure: { code: 'missing-target', proposalId: 'empty' },
    });
    expect(state(missing)).toEqual(unchanged);
  });

  it('refuses a round whose occurrences expand past the batch limit', async () => {
    const session = await open();
    const unchanged = state(session);
    expect(propose(session, replace('many', '00000005', 'a', 'b', 'all'))).toMatchObject({
      ok: false,
      failure: { code: 'limit-exceeded' },
    });
    expect(state(session)).toEqual(unchanged);
    snapshotOf(propose(session, replace('one', '00000005', 'a', 'b', 200)));
    expect(texts(session, 'accepted')[5]!.endsWith('a b')).toBe(true);

    const crowded = await open(fixture(paragraph('00000001', run('a'.repeat(10_001)))));
    const untouched = state(crowded);
    for (const occurrence of ['first', 1, 'all'] as const) {
      expect(propose(crowded, replace('crowded', '00000001', 'a', 'b', occurrence))).toMatchObject({
        ok: false,
        failure: { code: 'limit-exceeded', proposalId: 'crowded' },
      });
    }
    expect(state(crowded)).toEqual(untouched);
  });

  it('proposes into table cells through their persisted paragraph id', async () => {
    const session = await open();
    const snapshot = snapshotOf(propose(session, replace('cell', '0000C001', 'value', 'text')));
    expect(snapshot.proposals[0]!.paragraph).toMatchObject({
      story: 'body:t0:r0c0',
      paraId: '0000C001',
    });
    expect(texts(session, 'accepted', 'body:t0:r0c0')).toEqual(['cell text']);
    expect(texts(session, 'original', 'body:t0:r0c0')).toEqual(['cell value']);
  });

  it('refuses ambiguous and missing paragraphs without applying the rest of the round', async () => {
    const session = await open();
    const unchanged = state(session);
    const good = replace('good', '00000003', 'Keep', 'Hold');
    expect(propose(session, good, replace('twin', '0000000D', 'twin', 'sibling'))).toMatchObject({
      ok: false,
      failure: { code: 'ambiguous-target', proposalId: 'twin' },
    });
    expect(propose(session, good, replace('gone', '0BADF00D', 'x', 'y'))).toMatchObject({
      ok: false,
      failure: { code: 'missing-target', proposalId: 'gone' },
    });
    expect(
      propose(session, good, {
        ...replace('foreign', '00000006', 'Tail', 'End'),
        paragraph: {
          kind: 'session',
          sessionId: 'another-session',
          story: 'body',
          paraId: '00000006',
        },
      })
    ).toMatchObject({ ok: false, failure: { code: 'unsupported', proposalId: 'foreign' } });
    expect(state(session)).toEqual(unchanged);
    expect(texts(session, 'accepted')[2]).toBe('Keep this sentence.');
  });

  it('refuses overlapping proposals and edits over pending tracked changes', async () => {
    const session = await open();
    const unchanged = state(session);
    expect(
      propose(
        session,
        replace('left', '00000003', 'Keep this', 'Hold that'),
        replace('right', '00000003', 'this sentence', 'the line')
      )
    ).toMatchObject({ ok: false, failure: { code: 'overlapping-steps', proposalId: 'right' } });
    expect(
      propose(
        session,
        replace('keep', '00000003', 'Keep ', 'Hold '),
        replace('this', '00000003', 'this', 'that')
      )
    ).toMatchObject({ ok: false, failure: { code: 'overlapping-steps', proposalId: 'this' } });
    expect(state(session)).toEqual(unchanged);
    snapshotOf(propose(session, replace('first', '00000003', 'this', 'that')));
    const afterFirst = state(session);
    expect(propose(session, replace('again', '00000003', 'that', 'those'))).toMatchObject({
      ok: false,
      failure: { code: 'tracked-revision-conflict', proposalId: 'again' },
    });
    expect(state(session)).toEqual(afterFirst);
  });

  it('refuses a stale version and keeps retries idempotent by id and edit', async () => {
    const session = await open();
    const stale = session.version();
    const proposal = replace('p1', '00000003', 'Keep', 'Hold');
    const events: DocxProposalSnapshot[] = [];
    session.onProposalChange((snapshot) => events.push(snapshot));
    snapshotOf(session.proposeChanges({ expectVersion: stale, proposals: [proposal] }));
    expect(events).toHaveLength(1);
    const applied = state(session);

    const retry = snapshotOf(
      session.proposeChanges({
        expectVersion: stale,
        proposals: [{ ...proposal, suggest: { author: 'Other', date: '2026-09-30T00:00:00Z' } }],
      })
    );
    expect(retry).toEqual(applied.proposals);
    expect(state(session)).toEqual(applied);
    expect(events).toHaveLength(1);

    expect(
      session.proposeChanges({
        expectVersion: stale,
        proposals: [replace('p2', '00000006', 'Tail', 'End')],
      })
    ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(propose(session, replace('p1', '00000006', 'Tail', 'End'))).toMatchObject({
      ok: false,
      failure: { code: 'proposal-id-conflict', proposalId: 'p1' },
    });
    expect(propose(session, replace('p1', '00000003', 'Keep', 'Hold', 2))).toMatchObject({
      ok: false,
      failure: { code: 'proposal-id-conflict', proposalId: 'p1' },
    });
    expect(
      propose(
        session,
        replace('dup', '00000006', 'Tail', 'End'),
        replace('dup', '00000006', 'Tail', 'Stop')
      )
    ).toMatchObject({ ok: false, failure: { code: 'proposal-id-conflict', proposalId: 'dup' } });
    expect(propose(session, replace('', '00000006', 'Tail', 'End'))).toMatchObject({
      ok: false,
      failure: { code: 'invalid-step' },
    });
    expect(state(session)).toEqual(applied);
    expect(events).toHaveLength(1);

    const mixed = snapshotOf(propose(session, proposal, replace('p2', '00000006', 'Tail', 'End')));
    expect(mixed.proposals.map((record) => record.id)).toEqual(['p1', 'p2']);
    expect(events).toHaveLength(2);
    expect(propose(session, replace('p1', '00000003', 'Keep', 'Hold', 1))).toMatchObject({
      ok: true,
    });
    expect(events).toHaveLength(2);
  });

  it('registers a round before update listeners run', async () => {
    const session = await open(fixture(paragraph('00000001', run('one two three'))));
    const nested: DocxProposalResult[] = [];
    const seen: string[][] = [];
    session.onUpdate(() => {
      seen.push(session.getProposals().proposals.map((record) => record.id));
      if (nested.length > 0) return;
      nested.push(propose(session, replace('p1', '00000001', 'three', 'third')));
    });
    const outer = snapshotOf(propose(session, replace('p1', '00000001', 'one', 'first')));
    expect(seen[0]).toEqual(['p1']);
    expect(nested[0]).toMatchObject({
      ok: false,
      failure: { code: 'proposal-id-conflict', proposalId: 'p1' },
    });
    expect(session.getProposals().proposals).toEqual(outer.proposals);
    expect(texts(session, 'accepted')).toEqual(['first two three']);

    const reopening = await open();
    let reopened = false;
    reopening.onUpdate(() => {
      if (reopened) return;
      reopened = true;
      reopening.beginOpening('next');
    });
    const round = snapshotOf(propose(reopening, replace('p1', '00000003', 'Keep', 'Hold')));
    expect(round.proposals.map((record) => record.id)).toEqual(['p1']);
    expect(reopening.getProposals().proposals).toEqual([]);
  });

  it('keeps proposals that update listeners make after a reopen', async () => {
    const source = await open();
    const fresh = await createYrsSession({ clientId: nextClientId++ });
    sessions.push(fresh);
    const reopen: Array<[YrsSession, () => void]> = [
      [source, () => source.beginOpening('next')],
      [fresh, () => fresh.loadState(source.encodeState())],
    ];
    for (const [index, [session, action]] of reopen.entries()) {
      const id = `listener-${index}`;
      const paraId = index === 0 ? '00000006' : '00000003';
      let proposed: DocxProposalResult | null = null;
      const detach = session.onUpdate(() => {
        if (proposed) return;
        const paragraph = {
          kind: 'session',
          sessionId: session.paragraphIdentities().sessionId,
          story: 'body',
          paraId,
        } as const;
        proposed = propose(session, {
          ...insert(id, paraId, 'end', '!'),
          paragraph,
        });
      });
      action();
      detach();
      expect(proposed).toMatchObject({ ok: true });
      expect(session.getProposals().proposals.map((record) => record.id)).toEqual([id]);
      expect(snapshotOf(decide(session, [{ id, state: 'rejected' }])).proposals).toEqual([
        expect.objectContaining({ id, state: 'rejected' }),
      ]);
    }
  });

  it('delivers the newest snapshot last when a listener decides again', async () => {
    const session = await open();
    snapshotOf(
      propose(
        session,
        replace('a', '00000003', 'Keep', 'Hold'),
        replace('b', '00000006', 'Tail', 'End')
      )
    );
    const first: number[] = [];
    const second: number[] = [];
    session.onProposalChange((snapshot) => {
      first.push(snapshot.previewVersion);
      if (snapshot.previewVersion === 1)
        snapshotOf(decide(session, [{ id: 'b', state: 'rejected' }]));
    });
    session.onProposalChange((snapshot) => second.push(snapshot.previewVersion));
    snapshotOf(decide(session, [{ id: 'a', state: 'accepted' }]));
    expect(first).toEqual([1, 2]);
    expect(second).toEqual([2]);
    expect(session.getProposals().proposals.map((record) => record.state)).toEqual([
      'accepted',
      'rejected',
    ]);
  });

  it('keeps the engine refusal for locked content', async () => {
    const locked = `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/><w:tag w:val="clause"/></w:sdtPr><w:sdtContent>${paragraph(
      '0000E001',
      run('Locked clause')
    )}</w:sdtContent></w:sdt>`;
    const session = await open(fixture(`${paragraph('00000001', run('Open'))}${locked}`));
    const unchanged = state(session);
    expect(propose(session, replace('locked', '0000E001', 'clause', 'term'))).toMatchObject({
      ok: false,
      failure: { code: 'locked-target', proposalId: 'locked' },
    });
    expect(state(session)).toEqual(unchanged);
  });

  it('records a no-op proposal without revisions', async () => {
    const session = await open();
    const before = session.version();
    const snapshot = snapshotOf(
      propose(
        session,
        replace('same', '00000006', 'Tail', 'Tail'),
        insert('empty', '00000003', 'end', '')
      )
    );
    expect(snapshot.version).toBe(before);
    expect(snapshot.proposals).toEqual([
      expect.objectContaining({ id: 'same', changed: false, revisionIds: [] }),
      expect.objectContaining({ id: 'empty', changed: false, revisionIds: [] }),
    ]);
    expect(session.listRevisions()).toEqual([]);
    const beside = snapshotOf(
      propose(
        session,
        replace('kept', '00000003', 'Keep ', 'Keep '),
        replace('this', '00000003', 'this', 'that')
      )
    );
    expect(beside.proposals.map((record) => record.changed)).toEqual([false, false, false, true]);
  });

  it('keeps proposal rounds out of undo history', async () => {
    const session = await open();
    session.beginUndoCapture();
    session.insertText({ story: 'body', paraId: '00000006', offset: 4 }, '.');
    session.addUndoBoundary();
    expect(session.canUndo()).toBe(true);
    snapshotOf(propose(session, replace('p1', '00000003', 'Keep', 'Hold')));
    expect(session.undo()).toBe(true);
    expect(texts(session, 'accepted').at(-1)).toBe('Tail');
    expect(texts(session, 'accepted')[2]).toBe('Hold this sentence.');
    expect(session.canUndo()).toBe(false);
    expect(session.canRedo()).toBe(true);
  });

  it('previews decisions without touching the document, its version or history', async () => {
    const session = await open();
    snapshotOf(
      propose(
        session,
        replace('a', '00000001', 'world', 'earth'),
        replace('b', '00000003', 'this', 'that')
      )
    );
    const proposed = state(session);
    const accepted = texts(session, 'accepted');
    const events: DocxProposalSnapshot[] = [];
    session.onProposalChange((snapshot) => events.push(snapshot));
    const revisionsOf = (id: string) =>
      proposed.proposals.proposals.find((record) => record.id === id)!.revisionIds;

    const decided = snapshotOf(
      decide(session, [
        { id: 'a', state: 'accepted' },
        { id: 'b', state: 'rejected' },
      ])
    );
    expect(decided.previewVersion).toBe(1);
    expect(decided.version).toBe(proposed.version);
    expect(decided.proposals.map((record) => record.state)).toEqual(['accepted', 'rejected']);
    expect(proposalRevisionPreview(decided)).toEqual(
      Object.fromEntries(
        [
          ...revisionsOf('a').map((id) => [id, 'accepted'] as const),
          ...revisionsOf('b').map((id) => [id, 'rejected'] as const),
        ].sort(([x], [y]) => (x < y ? -1 : 1))
      )
    );
    expect(events.map((event) => event.previewVersion)).toEqual([1]);
    expect(session.listRevisions()).toEqual(proposed.revisions);
    expect(texts(session, 'accepted')).toEqual(accepted);
    expect([session.canUndo(), session.canRedo()]).toEqual([proposed.canUndo, proposed.canRedo]);

    expect(snapshotOf(decide(session, [{ id: 'a', state: 'accepted' }])).previewVersion).toBe(1);
    expect(events).toHaveLength(1);

    const undone = snapshotOf(decide(session, [{ id: 'a', state: 'proposed' }]));
    expect(undone.previewVersion).toBe(2);
    expect(proposalRevisionPreview(undone)).toEqual(
      Object.fromEntries(revisionsOf('b').map((id) => [id, 'rejected']))
    );
    const reset = snapshotOf(decide(session, [{ id: 'b', state: 'proposed' }]));
    expect(proposalRevisionPreview(reset)).toBeUndefined();
    expect(session.version()).toBe(proposed.version);
    expect(session.listRevisions()).toEqual(proposed.revisions);
  });

  it('refuses stale, unknown and malformed decisions without changing the preview', async () => {
    const session = await open();
    snapshotOf(propose(session, replace('a', '00000003', 'this', 'that')));
    snapshotOf(decide(session, [{ id: 'a', state: 'accepted' }]));
    const unchanged = state(session);
    const request = {
      expectVersion: session.version(),
      expectPreviewVersion: 1,
      changes: [{ id: 'a', state: 'rejected' as const }],
    };
    expect(session.setProposalStates({ ...request, expectPreviewVersion: 0 })).toMatchObject({
      ok: false,
      failure: { code: 'stale-preview' },
    });
    expect(session.setProposalStates({ ...request, expectVersion: 'stale' })).toMatchObject({
      ok: false,
      failure: { code: 'stale-version' },
    });
    expect(
      session.setProposalStates({
        ...request,
        changes: [
          { id: 'a', state: 'rejected' },
          { id: 'zzz', state: 'rejected' },
        ],
      })
    ).toMatchObject({ ok: false, failure: { code: 'unknown-proposal', proposalId: 'zzz' } });
    expect(
      session.setProposalStates({ ...request, changes: [{ id: 'a', state: 'done' as never }] })
    ).toMatchObject({ ok: false, failure: { code: 'invalid-step', proposalId: 'a' } });
    expect(
      session.setProposalStates({
        ...request,
        changes: [
          { id: 'a', state: 'rejected' },
          { id: 'a', state: 'proposed' },
        ],
      })
    ).toMatchObject({ ok: false, failure: { code: 'invalid-step', proposalId: 'a' } });
    expect(state(session)).toEqual(unchanged);
    expect(() => session.setProposalStates({ ...request, changes: null as never })).toThrow(
      TypeError
    );
    expect(() =>
      session.proposeChanges({
        expectVersion: session.version(),
        proposals: [{ id: 'x' } as never],
      })
    ).toThrow(TypeError);
  });

  it('fills an empty paragraph through an empty search in its mark formatting', async () => {
    const session = await open(fixture(FILLABLE));
    const snapshot = snapshotOf(
      propose(
        session,
        replace('cell', '0000F001', '', 'Cell fill'),
        replace('body', '0000F002', '', 'Body fill', 1)
      )
    );
    const story = snapshot.proposals[0]!.paragraph.story;
    expect(story).toBe('body:t0:r0c0');
    expect(texts(session, 'accepted', story)).toEqual(['Cell fill']);
    expect(texts(session, 'original', story)).toEqual(['']);
    expect(texts(session, 'accepted')[0]).toBe('Body fill');
    const segment = (story: string, text: string) =>
      session
        .storySegments(story)
        .find((entry) => entry.kind === 'text' && entry.text === text)!.attributes;
    expect(segment(story, 'Cell fill')).toMatchObject({
      fontSize: { size: 18 },
      ins: { author: 'Atira' },
    });
    expect(segment('body', 'Body fill').fontSize ?? null).toBeNull();
    const xml = new TextDecoder().decode(
      unzipContainer((await saveYrsDocx(session)).bytes)['word/document.xml']
    );
    const filled = xml.slice(0, xml.indexOf('>Cell fill<'));
    const run = filled.slice(Math.max(filled.lastIndexOf('<w:r>'), filled.lastIndexOf('<w:r ')));
    expect(run).toContain('<w:sz w:val="18"/>');
  });

  it('fills only an empty paragraph, at its first occurrence', async () => {
    const session = await open(fixture(FILLABLE));
    const unchanged = state(session);
    for (const paraId of ['0000F003', '0000F004', '0000F005', '0000F006', '0000F008']) {
      expect(propose(session, replace(paraId, paraId, '', 'x'))).toMatchObject({
        ok: false,
        failure: { code: 'missing-target', proposalId: paraId },
      });
    }
    expect(propose(session, replace('second', '0000F002', '', 'x', 2))).toMatchObject({
      ok: false,
      failure: { code: 'missing-target', proposalId: 'second' },
    });
    expect(propose(session, replace('twin', '0000F007', '', 'x'))).toMatchObject({
      ok: false,
      failure: { code: 'ambiguous-target', proposalId: 'twin' },
    });
    expect(propose(session, replace('struck', '0000F009', '', 'x'))).toMatchObject({
      ok: false,
      failure: { code: 'tracked-revision-conflict', proposalId: 'struck' },
    });
    expect(
      propose(
        session,
        replace('one', '0000F002', '', 'x'),
        replace('two', '0000F002', '', 'y')
      )
    ).toMatchObject({ ok: false, failure: { code: 'overlapping-steps' } });
    expect(state(session)).toEqual(unchanged);
  });

  it('withdraws proposals as their decisions preview them, outside undo history', async () => {
    const session = await open();
    snapshotOf(
      propose(
        session,
        replace('rejected', '00000003', 'this', 'that'),
        replace('accepted', '00000006', 'Tail', 'End'),
        insert('undecided', '00000001', 'end', '!'),
        replace('kept', '00000002', 'and', 'or'),
        replace('cell', '0000C001', 'value', 'text')
      )
    );
    snapshotOf(
      decide(session, [
        { id: 'rejected', state: 'rejected' },
        { id: 'accepted', state: 'accepted' },
        { id: 'cell', state: 'rejected' },
      ])
    );
    expect(propose(session, replace('again', '00000003', 'this', 'the'))).toMatchObject({
      ok: false,
      failure: { code: 'missing-target' },
    });
    await saveYrsDocx(session);
    const events: DocxProposalSnapshot[] = [];
    session.onProposalChange((snapshot) => events.push(snapshot));
    const before = session.version();
    const withdrawn = snapshotOf(
      session.withdrawProposals({
        expectVersion: before,
        ids: ['rejected', 'accepted', 'undecided', 'cell', 'unknown'],
      })
    );
    expect(withdrawn.version).not.toBe(before);
    expect(withdrawn.previewVersion).toBe(2);
    expect(withdrawn.proposals.map((record) => record.id)).toEqual(['kept']);
    expect(events.map((event) => event.proposals.length)).toEqual([1]);
    expect(texts(session, 'original')).toEqual([
      'Hello world',
      'aaaa and aaaa',
      'Keep this sentence.',
      'First twin',
      'Second twin',
      'a '.repeat(200).trimEnd(),
      'End',
    ]);
    expect(texts(session, 'accepted')[1]).toBe('aaaa or aaaa');
    expect(texts(session, 'accepted', 'body:t0:r0c0')).toEqual(['cell value']);
    expect(new Set(session.listRevisions().map((revision) => revision.revisionId))).toEqual(
      new Set(withdrawn.proposals[0]!.revisionIds)
    );
    expect(session.canUndo()).toBe(false);
    const saved = new TextDecoder().decode(
      unzipContainer((await saveYrsDocx(session)).bytes)['word/document.xml']
    );
    expect(saved.match(/<w:(ins|del) /g)).toHaveLength(2);
    expect(saved).not.toContain('>that<');
    expect(saved).not.toContain('>Tail<');
    expect(saved).not.toContain('>text<');

    snapshotOf(propose(session, replace('again', '00000003', 'this', 'the')));
    expect(texts(session, 'accepted')[2]).toBe('Keep the sentence.');
    const retried = session.getProposals();
    expect(
      session.withdrawProposals({ expectVersion: before, ids: ['rejected', 'accepted'] })
    ).toEqual({ ok: true, snapshot: retried });
    expect(
      session.withdrawProposals({ expectVersion: before, ids: ['again'] })
    ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    expect(session.getProposals()).toEqual(retried);
    expect(() => session.withdrawProposals({ expectVersion: before } as never)).toThrow(TypeError);
  });

  it('refuses to withdraw a proposal whose revision also marks a paragraph change', async () => {
    const session = await open();
    const [record] = snapshotOf(
      propose(session, replace('styled', '00000003', 'this', 'that'))
    ).proposals;
    const mark = { paraId: '00000003', offset: 0 };
    session.setParagraphAttrs(
      { story: 'body', start: mark, end: mark },
      { alignment: 'right' },
      { name: SUGGEST.author, date: SUGGEST.date }
    );
    const changed = session.listRevisions().find((revision) => revision.kind === 'pPrChange');
    expect(record!.revisionIds).toContain(changed!.revisionId);
    const unchanged = state(session);
    expect(
      session.withdrawProposals({ expectVersion: session.version(), ids: ['styled'] })
    ).toMatchObject({
      ok: false,
      failure: { code: 'tracked-revision-conflict', proposalId: 'styled' },
    });
    expect(state(session)).toEqual(unchanged);
  });

  it('settles a revision two withdrawn proposals share as the preview shows it', async () => {
    const session = await open();
    const settled: (readonly string[])[][] = [];
    const registry = createProposalRegistry({
      version: () => session.version(),
      resolveParagraphAnchor: (anchor) => session.resolveParagraphAnchor(anchor),
      findText: (request) => session.findText(request),
      readParagraphs: (request) => session.readParagraphs(request),
      applyEdits: (request) => {
        const result = session.applyEdits(request);
        if (!result.ok) return result;
        return {
          ...result,
          receipts: result.receipts.map((receipt) => ({ ...receipt, revisionIds: ['shared'] })),
        };
      },
      listRevisions: () => [],
      settleRevisions: (accept, reject) => settled.push([accept, reject]),
    });
    for (const [id, paraId] of [
      ['first', '00000003'],
      ['second', '00000006'],
    ] as const) {
      snapshotOf(
        registry.propose({
          expectVersion: session.version(),
          proposals: [insert(id, paraId, 'end', '!')],
        })
      );
    }
    const decided = snapshotOf(
      registry.setStates({
        expectVersion: session.version(),
        expectPreviewVersion: registry.snapshot().previewVersion,
        changes: [
          { id: 'first', state: 'accepted' },
          { id: 'second', state: 'rejected' },
        ],
      })
    );
    expect(proposalRevisionPreview(decided)).toEqual({ shared: 'rejected' });

    snapshotOf(registry.withdraw({ expectVersion: session.version(), ids: ['first', 'second'] }));
    expect(settled).toEqual([[[], ['shared']]]);
  });

  it('forgets proposals when the session opens another document', async () => {
    const session = await open();
    snapshotOf(propose(session, replace('a', '00000003', 'this', 'that')));
    snapshotOf(decide(session, [{ id: 'a', state: 'accepted' }]));
    const events: DocxProposalSnapshot[] = [];
    session.onProposalChange((snapshot) => events.push(snapshot));
    session.openDocx(fixture(), false);
    expect(session.getProposals()).toEqual({
      version: session.version(),
      previewVersion: 2,
      proposals: [],
    });
    expect(events).toHaveLength(1);
    expect(events[0]!.proposals).toEqual([]);
    snapshotOf(propose(session, replace('a', '00000006', 'Tail', 'End')));
    expect(session.getProposals().proposals.map((record) => record.id)).toEqual(['a']);
  });
});
