import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm, type EditSession } from '../wasm/edit';
import { EditSession as WasmEditSession } from '../wasm/generated/edit/docx_edit';
import { createYrsSession, type YrsRawOp, type YrsSession } from './index';
import { computeProposalGeometryMirror, type ProposalGeometryReader } from './proposalGeometry';
import type { DocxProposalSnapshot } from './proposals';
import {
  createResidentEngineSession,
  readLegacyProposalRevisions,
  type ResidentEngineSession,
} from './residentEngineSession';
import { readResidentSearch } from './residentSearch';
import {
  createYrsLocProjectionFromOutline,
  createYrsPositionProjection,
  yrsLocToProjectedDisplayPosition,
} from './yrsPositionProjection';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OWNED = ['owned', 'other'];
const p = (id: string, content: string) => `<w:p w14:paraId="${id}">${content}</w:p>`;
const r = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const cell = (content: string, properties = '') => `<w:tc><w:tcPr>${properties}</w:tcPr>${content}</w:tc>`;
const table = (rows: string) => `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows}</w:tbl>`;

function docx(body: string): Uint8Array {
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr/></w:body></w:document>`)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

interface Segment {
  text?: string;
  kind?: string;
  payload?: Record<string, unknown>;
  attrs?: Record<string, unknown>;
}

function story(main: YrsSession, id: string, segments: Segment[]): void {
  main.createStory(id, '');
  let index = 0;
  const ops: YrsRawOp[] = [{ op: 'delete', index: 0, len: 1 }];
  for (const segment of segments) {
    if (segment.text !== undefined) {
      ops.push({ op: 'insert', index, text: segment.text, attrs: segment.attrs });
      index += segment.text.length;
    } else {
      ops.push({ op: 'insertEmbed', index, kind: segment.kind!, payload: segment.payload, attrs: segment.attrs });
      index += 1;
    }
  }
  main.applySeedRawOps(id, ops);
}

const pilcrow = (paraId: string): Segment => ({ kind: 'pilcrow', payload: { paraId } });
const text = (value: string, attrs?: Record<string, unknown>): Segment => ({ text: value, attrs });
const ins = { ins: { id: 'owned' } };

async function replica(main?: YrsSession): Promise<{ resident: ResidentEngineSession; raw: EditSession }> {
  const resident = await createResidentEngineSession();
  let raw: EditSession | undefined;
  const version = WasmEditSession.prototype.version;
  const capture = spyOn(WasmEditSession.prototype, 'version').mockImplementation(function (this: EditSession) {
    raw = this;
    return version.call(this);
  });
  try {
    resident.geometryReader.version();
  } finally {
    capture.mockRestore();
  }
  try {
    expect(typeof raw!.geometry_position_outline_json).toBe('function');
    expect(typeof raw!.proposal_revision_ranges_json).toBe('function');
    if (main) resident.loadState(main.encodeState());
    return { resident, raw: raw! };
  } catch (error) {
    resident.destroy();
    throw error;
  }
}

function same(actual: unknown, expected: unknown): void {
  expect(actual).toEqual(expected);
  expect(JSON.stringify(actual)).toBe(JSON.stringify(expected));
}

function snapshot(resident: ResidentEngineSession, ids: readonly string[]): DocxProposalSnapshot {
  const identities = resident.paragraphIdentities();
  return {
    version: resident.geometryReader.version(),
    previewVersion: 0,
    proposals: identities.paragraphs.filter(({ session }) => session !== null).map(({ session }, index) => ({
      id: `p${index}`,
      state: index === 0 ? 'rejected' : 'proposed',
      paragraph: session!,
      revisionIds: index === 0 ? ids : [],
      changed: index === 0,
    })),
  };
}

function parity(
  resident: ResidentEngineSession,
  raw: EditSession,
  ids: readonly string[] = OWNED,
  current = snapshot(resident, ids),
  navigation = true
): void {
  const reader = resident.geometryReader;
  const legacy: ProposalGeometryReader = {
    ...reader,
    positionOutline: undefined,
    proposalRevisions: (owned) => readLegacyProposalRevisions(reader, owned),
  };
  same(reader.proposalRevisions!(ids), legacy.proposalRevisions!(ids));
  const native = JSON.parse(raw.proposal_revision_ranges_json(JSON.stringify(ids)));
  if (Array.isArray(native)) same(native, readLegacyProposalRevisions(reader, ids));
  for (const includeNavigation of navigation ? [false, true] : [false]) {
    same(computeProposalGeometryMirror(reader, current, includeNavigation),
      computeProposalGeometryMirror(legacy, current, includeNavigation));
  }
  for (const story of reader.storyIds()) {
    const spans: { paraId: string; length: number }[] = [];
    let length = 0;
    for (const segment of reader.storySegments(story)) {
      if (segment.kind === 'pilcrow') {
        spans.push({ paraId: segment.paraId, length });
        length = 0;
      } else {
        length += segment.kind === 'text' ? segment.text.length : 1;
      }
    }
    expect(reader.paragraphSpans(story)).toEqual(spans);
  }
  const outline = reader.positionOutline!('body');
  const old = createYrsPositionProjection(reader, 'body');
  const fast = outline ? createYrsLocProjectionFromOutline(outline) : old;
  for (const story of reader.storyIds()) {
    for (const { paraId, length } of reader.paragraphSpans(story)) {
      for (const offset of [-Infinity, NaN, ...Array.from({ length: length + 5 }, (_, index) => index - 2), length + 0.5, Infinity]) {
        const loc = { story, paraId, offset };
        expect(fast?.positionForLoc(loc)).toBe(old?.positionForLoc(loc));
        expect(yrsLocToProjectedDisplayPosition(reader, () => fast, loc)).toBe(
          yrsLocToProjectedDisplayPosition(reader, () => old, loc)
        );
      }
    }
    const missing = { story, paraId: 'missing', offset: 3 };
    expect(yrsLocToProjectedDisplayPosition(reader, () => fast, missing)).toBe(
      yrsLocToProjectedDisplayPosition(reader, () => old, missing)
    );
  }
  const searchReader = {
    ...reader,
    searchText: resident.searchText,
    resolveStickyPosition: resident.resolveStickyPosition,
  };
  const oldSearchReader = { ...searchReader, positionOutline: undefined };
  for (const query of ['cat', '😀', 'cell', 'inside', 'X', '']) {
    for (const caseSensitive of [false, true]) {
      same(readResidentSearch(searchReader, query, caseSensitive),
        readResidentSearch(oldSearchReader, query, caseSensitive));
      const hit = resident.searchText(query)[0];
      if (hit) {
        const carry = resident.encodeStickyPosition({ story: hit.story, paraId: hit.paraId, offset: hit.start });
        same(readResidentSearch(searchReader, query, caseSensitive, carry),
          readResidentSearch(oldSearchReader, query, caseSensitive, carry));
      }
    }
  }
  const gate = { ...reader, hasStory: () => false };
  expect(yrsLocToProjectedDisplayPosition(gate, () => fast, { story: 'body', paraId: 'missing', offset: 0 })).toBeNull();
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

describe('native proposal geometry outline parity', () => {
  const fixtures: Array<[string, Uint8Array]> = [
    ['table and post-table paragraphs', docx(p('00000001', r('cat 😀cat')) + table(`<w:tr>${cell(p('0000C001', r('cell cat')))}</w:tr>`) + p('00000002', r('cat tail')))],
    ['merged cells, nested tables and empty rows/cells', docx(
      table(`<w:tr>${cell(p('0000C001', r('cell cat')), '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>')}${cell(table(`<w:tr>${cell(p('0000C002', r('nested cat')))}</w:tr>`))}</w:tr><w:tr>${cell(p('0000C003', ''), '<w:gridSpan w:val="2"/><w:vMerge/>')}${cell(p('0000C004', ''))}</w:tr><w:tr/>`) + p('00000001', r('tail cat'))
    )],
    ['block SDT', docx(p('00000001', r('cat')) + `<w:sdt><w:sdtPr><w:alias w:val="Control"/><w:tag w:val="control"/></w:sdtPr><w:sdtContent>${p('0000C001', r('inside cat'))}</w:sdtContent></w:sdt>` + p('00000002', r('cat tail')))],
    ['leading and inline page/column breaks', docx(p('00000001', '<w:r><w:br w:type="page"/><w:br w:type="column"/><w:t>cat</w:t><w:br w:type="page"/><w:t>😀cat</w:t><w:br w:type="column"/></w:r>') + p('00000002', r('cat tail')))],
    ['content control template', new Uint8Array(readFileSync(resolve(import.meta.dir, '__fixtures__/content-controls/template.docx')))],
  ];

  for (const [name, bytes] of fixtures) {
    test(name, async () => {
      const main = await createYrsSession({ clientId: 84101 });
      let resident: ResidentEngineSession | undefined;
      try {
        main.openDocx(bytes, true);
        const target = main.paragraphs('body').find(({ text }) => text.includes('cat'));
        let current: DocxProposalSnapshot | undefined;
        if (target) {
          const result = main.proposeChanges({
            expectVersion: main.version(),
            proposals: [{ id: 'edit', paragraph: { kind: 'session', sessionId: main.paragraphIdentities().sessionId, story: 'body', paraId: target.paraId }, suggest: { author: 'Reviewer', date: '2026-09-29T12:00:00Z' }, op: 'replaceText', search: 'cat', replaceWith: 'Xcat', occurrence: 'all' }],
          });
          if (!result.ok) throw new Error(result.failure.message);
          current = result.snapshot;
        }
        const loaded = await replica(main);
        resident = loaded.resident;
        expect(resident.geometryReader.positionOutline!('body')).not.toBeNull();
        const ids = current?.proposals.flatMap(({ revisionIds }) => revisionIds) ?? OWNED;
        parity(resident, loaded.raw, ids, current);
        if (current) {
          parity(resident, loaded.raw, ids, { ...current, previewVersion: 1, proposals: current.proposals.map((proposal) => ({ ...proposal, state: 'accepted' })) });
          parity(resident, loaded.raw, ids, { ...current, proposals: [] });
        }
        const nested = main.storyIds().find((id) => id.startsWith('body:') && main.paragraphs(id).length > 0);
        if (nested) {
          const vector = resident.encodeStateVector();
          const paraId = main.paragraphs(nested)[0]!.paraId;
          main.insertText({ story: nested, paraId, offset: 0 }, '😀cat');
          resident.applyUpdate(main.encodeStateAsUpdate(vector));
          parity(resident, loaded.raw, ids);
          resident.loadState(main.encodeState());
          parity(resident, loaded.raw, ids);
        }
      } finally {
        resident?.destroy();
        main.destroy();
      }
    });
  }

  test('cold resident open and table merges refresh positions', async () => {
    const loaded = await replica();
    try {
      loaded.resident.openDocx(fixtures[0]![1]);
      parity(loaded.resident, loaded.raw);
    } finally {
      loaded.resident.destroy();
    }
    const main = await createYrsSession({ clientId: 84108 });
    let resident: ResidentEngineSession | undefined;
    try {
      const { paraId } = main.createStory('body', 'cat tail');
      const inserted = main.insertTable({ story: 'body', paraId, offset: 0 }, 2, 2);
      for (const id of inserted.createdStoryIds) {
        main.insertText({ story: id, paraId: main.paragraphs(id)[0]!.paraId, offset: 0 }, 'cell cat');
      }
      const loaded = await replica(main);
      resident = loaded.resident;
      parity(resident, loaded.raw);
      for (const row of [0, 1]) {
        const vector = resident.encodeStateVector();
        main.mergeCells({
          anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 },
          head: { story: 'body', tableIndex: 0, row, column: 1 },
        });
        resident.applyUpdate(main.encodeStateAsUpdate(vector));
        parity(resident, loaded.raw);
      }
    } finally {
      resident?.destroy();
      main.destroy();
    }
  });

  test('raw block embeds after inline text, default cell indices, shared/cyclic refs and unterminated content', async () => {
    const main = await createYrsSession({ clientId: 84102 });
    let resident: ResidentEngineSession | undefined;
    try {
      story(main, 'body', [
        { kind: 'pageBreak' }, { kind: 'columnBreak' }, text('😀cat'),
        { kind: 'table', payload: { rows: [{ cells: [{ tcPr: { colspan: 2, rowspan: 2 } }, {}] }, { cells: [] }] } },
        { kind: 'blockSdt', payload: { story: 'body:shared' } },
        { kind: 'blockSdt', payload: { story: 'body:shared' } },
        { kind: 'field', payload: { text: 'long field display text' } }, { kind: 'image' },
        { kind: 'noteRef' }, { kind: 'inlineSdt' }, { kind: 'pageBreak' }, { kind: 'columnBreak' },
        pilcrow('p'), text('unterminated cat'),
      ]);
      story(main, 'body:t0:r0c0', [pilcrow('empty')]);
      story(main, 'body:t0:r0c1', [{ kind: 'table', payload: { rows: [] } }, text('cell cat'), pilcrow('cp')]);
      story(main, 'body:shared', [{ kind: 'blockSdt', payload: { story: 'body' } }, text('inside cat'), pilcrow('sp')]);
      story(main, 'body:orphan', [text('cat'), pilcrow('orphan')]);
      const loaded = await replica(main);
      resident = loaded.resident;
      expect(resident.geometryReader.positionOutline!('body')).not.toBeNull();
      parity(resident, loaded.raw, OWNED, undefined, false);
    } finally {
      resident?.destroy();
      main.destroy();
    }
  });

  test('revision coalescing, stable ins/del ties, malformed stamps and UTF-16 offsets', async () => {
    const main = await createYrsSession({ clientId: 84103 });
    let resident: ResidentEngineSession | undefined;
    try {
      story(main, 'body', [
        text('😀', { ...ins, del: { info: { revisionId: 'other' } } }),
        { kind: 'image', attrs: { ...ins, del: { id: 'other' } } },
        text('x', { ins: { id: 'unowned' }, del: { id: 'other' } }),
        text('cat', { ins: { id: 'ignored', info: { id: 'owned' } }, del: { id: 'other' } }),
        pilcrow('p1'), text('cat', { ins: { revisionId: 'owned' }, del: { id: null, revisionId: 'other' } }),
        text('cat', { ins: { id: 'owned' }, del: [] }), pilcrow('p2'),
        text('cat', { ins: [], del: { info: [], revisionId: 'other' } }), pilcrow('p3'),
      ]);
      story(main, 'hf:header', [text('cat'), pilcrow('hp')]);
      story(main, 'fn:1', [text('cat'), pilcrow('np')]);
      story(main, 'body:orphan', [text('cat', ins), pilcrow('op')]);
      const loaded = await replica(main);
      resident = loaded.resident;
      expect(JSON.parse(loaded.raw.proposal_revision_ranges_json(JSON.stringify(OWNED)))).toBeArray();
      parity(resident, loaded.raw);
    } finally {
      resident?.destroy();
      main.destroy();
    }
  });

  for (const key of ['id', 'revisionId', 'pPrIns', 'pPrDel', 'pPrChange', 'trIns', 'trDel', 'tableIns', 'tableDel']) {
    for (const kind of ['pilcrow', 'image']) {
      test(`fallback for ${key} nested in ${kind}`, async () => {
        const main = await createYrsSession({ clientId: 84104 });
        let resident: ResidentEngineSession | undefined;
        try {
          story(main, 'body', [text('cat', ins), { kind, payload: { paraId: 'p', extra: [{ nested: { [key]: null } }] } }, ...(kind === 'pilcrow' ? [] : [pilcrow('p')])]);
          const loaded = await replica(main);
          resident = loaded.resident;
          expect(JSON.parse(loaded.raw.proposal_revision_ranges_json(JSON.stringify(OWNED)))).toBe('fallback');
          parity(resident, loaded.raw);
        } finally {
          resident?.destroy();
          main.destroy();
        }
      });
    }
  }

  test('fallback for duplicate paragraph IDs and owned unterminated tails', async () => {
    for (const duplicate of [false, true]) {
      const main = await createYrsSession({ clientId: 84105 });
      let resident: ResidentEngineSession | undefined;
      try {
        story(main, 'body', [text('cat', ins), pilcrow('p'), text('cat', ins), ...(duplicate ? [pilcrow('q')] : [])]);
        const loaded = await replica(main);
        resident = loaded.resident;
        if (duplicate) {
          loaded.raw.apply_seed_raw_ops('body', JSON.stringify([{ op: 'setEmbedAttr', index: 7, key: 'paraId', value: 'p' }]));
          expect(resident.geometryReader.paragraphSpans('body').map(({ paraId }) => paraId)).toEqual(['p', 'p']);
        }
        expect(JSON.parse(loaded.raw.proposal_revision_ranges_json(JSON.stringify(OWNED)))).toBe('fallback');
        parity(resident, loaded.raw);
      } finally {
        resident?.destroy();
        main.destroy();
      }
    }
  });

  test('legacy for numeric revision IDs and non-ASCII story sorting', async () => {
    for (const numeric of [false, true]) {
      const main = await createYrsSession({ clientId: 84106 });
      let resident: ResidentEngineSession | undefined;
      try {
        story(main, 'body', [text('cat', numeric ? { ins: { id: -0 }, del: { id: 1e21 } } : ins), pilcrow('p')]);
        if (!numeric) {
          story(main, '😀', [text('cat', ins), pilcrow('np')]);
          story(main, '\uE000', [text('cat', ins), pilcrow('bp')]);
        }
        const loaded = await replica(main);
        resident = loaded.resident;
        expect(JSON.parse(loaded.raw.proposal_revision_ranges_json(JSON.stringify(OWNED)))).toBe('legacy');
        parity(resident, loaded.raw, numeric ? ['0', '1e+21'] : OWNED);
      } finally {
        resident?.destroy();
        main.destroy();
      }
    }
  });

  test('legacy position outline for coercible non-string story refs and primitive row/cell shapes', async () => {
    for (const embed of [
      { kind: 'blockSdt', payload: { story: 42 } },
      { kind: 'table', payload: { rows: [7] } },
      { kind: 'table', payload: { rows: [{ cells: [false] }] } },
    ]) {
      const main = await createYrsSession({ clientId: 84107 });
      let resident: ResidentEngineSession | undefined;
      try {
        story(main, 'body', [embed, text('cat', ins), pilcrow('p')]);
        story(main, '42', [text('inside cat'), pilcrow('cp')]);
        story(main, 'body:t0:r0c0', [text('cell cat'), pilcrow('dp')]);
        const loaded = await replica(main);
        resident = loaded.resident;
        expect(resident.geometryReader.positionOutline!('body')).toBeNull();
        parity(resident, loaded.raw, OWNED, undefined, false);
      } finally {
        resident?.destroy();
        main.destroy();
      }
    }
  });
});
