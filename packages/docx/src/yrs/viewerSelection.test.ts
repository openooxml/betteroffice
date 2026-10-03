import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { DisplayPositionIndex } from './displayPositionIndex';
import { resolveYrsPointPosition } from './pointPosition';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import { createProposalRegistry } from './proposals';
import {
  resolveBookmarkPosition,
  resolveSelectionText,
  resolveSelectionUnit,
} from './viewerSelection';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const p = (id: string, content: string) => `<w:p w14:paraId="${id}">${content}</w:p>`;
const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const cell = (content: string) => `<w:tc><w:tcPr/>${content}</w:tc>`;

function docx(body: string): Uint8Array {
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr/></w:body></w:document>`)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const BODY =
  p('00000001', r('Alpha beta gamma.')) +
  p('00000002', `<w:bookmarkStart w:id="0" w:name="target"/>${r('Second ')}<w:del w:id="1" w:author="A" w:date="2026-10-01T00:00:00Z"><w:r><w:delText>gone </w:delText></w:r></w:del>${r('line')}<w:bookmarkEnd w:id="0"/>`) +
  `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
  `<w:tr>${cell(p('00000003', r('A1')))}${cell(p('00000004', r('B1')))}</w:tr>` +
  `<w:tr>${cell(p('00000005', r('A2')))}${cell(p('00000006', r('B2')))}</w:tr></w:tbl>` +
  p('00000007', r('After the table.'));

let resident: ResidentEngineSession;
let index: DisplayPositionIndex;

beforeAll(async () => {
  preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  resident = await createResidentEngineSession();
  resident.openDocx(docx(BODY));
  index = new DisplayPositionIndex({
    ...resident.geometryReader,
    selectionText: resident.selectionText,
  });
});

afterAll(() => resident.destroy());

async function withDocument(body: string, check: (resident: ResidentEngineSession, index: DisplayPositionIndex) => void) {
  const resident = await createResidentEngineSession();
  try {
    resident.openDocx(docx(body));
    const index = new DisplayPositionIndex({
      ...resident.geometryReader,
      selectionText: resident.selectionText,
    });
    check(resident, index);
  } finally {
    resident.destroy();
  }
}

function cellStory(paraId: string): string {
  for (const story of resident.geometryReader.storyIds()) {
    if (resident.geometryReader.paragraphs(story).some((paragraph) => paragraph.paraId === paraId)) {
      return story;
    }
  }
  throw new Error(`no story holds ${paraId}`);
}

/** Units of `paraId` before its text: block embeds such as a table ahead of it count in its offsets. */
function lead(paraId: string, text: string): number {
  const span = resident.geometryReader
    .paragraphSpans(cellStory(paraId))
    .find((candidate) => candidate.paraId === paraId);
  return span!.length - text.length;
}

function at(paraId: string, offset: number): number {
  const position = index.positionOf({ story: cellStory(paraId), paraId, offset }, 'body');
  if (position === null) throw new Error(`no display position for ${paraId}:${offset}`);
  return position;
}

describe('viewer selection reads', () => {
  test('a proposal decision without a document version change preserves display positions and selection text', async () => {
    await withDocument(p('00000001', r('Alpha beta')), (resident, index) => {
      const registry = createProposalRegistry(resident.proposalEngine);
      const proposed = registry.propose({ expectVersion: resident.geometryReader.version(), proposals: [{
        id: 'replace',
        paragraph: { kind: 'persisted', story: { kind: 'body', partUri: '/word/document.xml' }, paraId: '00000001' },
        suggest: { author: 'A', date: '2026-10-01T00:00:00Z' },
        op: 'replaceText', search: 'Alpha', replaceWith: 'Omega',
      }] });
      if (!proposed.ok) throw new Error(proposed.failure.message);
      const version = resident.geometryReader.version();
      const all = resolveSelectionUnit(index, 'body', 0, 'story', version)!;
      const before = resolveSelectionText(index, 'body', all.anchor, all.head, version);
      const decided = registry.setStates({
        expectVersion: version, expectPreviewVersion: proposed.snapshot.previewVersion,
        changes: [{ id: 'replace', state: 'rejected' }],
      });
      expect(decided.ok).toBe(true);
      expect(resident.geometryReader.version()).toBe(version);
      expect(resolveSelectionUnit(index, 'body', 0, 'story', version)).toEqual(all);
      expect(resolveSelectionText(index, 'body', all.anchor, all.head, version)?.text).toBe(before?.text);
    });
  });

  test('a word and a paragraph expand around a display position', () => {
    const version = resident.geometryReader.version();
    expect(resolveSelectionUnit(index, 'body', at('00000001', 7), 'word', version)).toEqual({
      anchor: at('00000001', 6),
      head: at('00000001', 10),
    });
    expect(resolveSelectionUnit(index, 'body', at('00000001', 7), 'paragraph', version)).toEqual({
      anchor: at('00000001', 0),
      head: at('00000001', 17),
    });
    expect(resolveSelectionUnit(index, 'body', at('00000004', 1), 'word', version)).toEqual({
      anchor: at('00000004', 0),
      head: at('00000004', 2),
    });
  });

  test('select-all spans the body story', () => {
    const version = resident.geometryReader.version();
    const all = resolveSelectionUnit(index, 'body', 0, 'story', version)!;
    expect(all.anchor).toBeLessThan(at('00000001', 0));
    expect(all.head).toBeGreaterThanOrEqual(at('00000007', lead('00000007', 'After the table.') + 16));
    expect(resolveSelectionText(index, 'body', all.anchor, all.head, version)?.text).toBe(
      'Alpha beta gamma.\nSecond gone line\nA1\tB1\nA2\tB2\nAfter the table.'
    );
  });

  test('select-all, a drag from the start and the table alone copy a leading table', async () => {
    const table = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>` +
      `<w:tr>${cell(p('00000003', r('A1')))}${cell(p('00000004', r('B1')))}</w:tr></w:tbl>`;
    await withDocument(table + p('00000007', r('Tail')), (resident, index) => {
      const version = resident.geometryReader.version();
      const all = resolveSelectionUnit(index, 'body', 0, 'story', version)!;
      expect(all.anchor).toBe(0);
      expect(resolveSelectionText(index, 'body', all.anchor, all.head, version)?.text).toBe('A1\tB1\nTail');
      const tail = index.positionOf({ story: 'body', paraId: '00000007', offset: 3 }, 'body')!;
      expect(resolveSelectionText(index, 'body', 0, tail, version)?.text).toBe('A1\tB1\nTa');
      const block = index.projection('body')!.tableAtStart(0)!;
      expect(resolveSelectionText(index, 'body', 0, block.nodeSize, version)?.text).toBe('A1\tB1\n');
    });
  });

  test('a range reads as clipboard text with its accepted-view range', () => {
    const version = resident.geometryReader.version();
    const text = resolveSelectionText(index, 'body', at('00000002', 13), at('00000001', 6), version);
    expect(text).toMatchObject({
      text: 'beta gamma.\nSecond gone l',
      range: {
        story: 'body',
        start: { paraId: '00000001', offset: 6 },
        end: { paraId: '00000002', offset: 8 },
        view: 'accepted',
      },
    });
  });

  test('cells of one table read as tab-separated rows, and a range into a table takes the whole table', () => {
    const version = resident.geometryReader.version();
    expect(resolveSelectionText(index, 'body', at('00000003', 1), at('00000006', 1), version)).toMatchObject({
      text: 'A1\tB1\nA2\tB2',
      range: null,
    });
    expect(resolveSelectionText(index, 'body', at('00000002', 7), at('00000004', 1), version)).toMatchObject({
      text: 'gone line\nA1\tB1\nA2\tB2\n',
      range: null,
    });
    const after = lead('00000007', 'After the table.');
    expect(resolveSelectionText(index, 'body', at('00000002', 7), at('00000007', after + 5), version)).toMatchObject({
      text: 'gone line\nA1\tB1\nA2\tB2\nAfter',
      range: {
        story: 'body',
        start: { paraId: '00000002', offset: 7 },
        end: { paraId: '00000007', offset: 5 },
        view: 'accepted',
      },
    });
  });

  test('a hit resolves to its accepted-view offset', () => {
    const version = resident.geometryReader.version();
    const hit = { position: at('00000002', 13), pageIndex: 0, region: 'body' as const };
    expect(resolveYrsPointPosition(index, hit, version)).toEqual({
      ...hit,
      version,
      target: {
        kind: 'range',
        story: 'body',
        start: { paraId: '00000002', offset: 8 },
        end: { paraId: '00000002', offset: 8 },
        view: 'accepted',
      },
    });
  });

  test('bookmarks resolve to their paragraph', () => {
    const version = resident.geometryReader.version();
    expect(resolveBookmarkPosition(index, 'body', 'target', version)).toBe(at('00000002', 0) - 1);
  });

  test('a read for another version answers null', () => {
    const stale = `${resident.geometryReader.version()}-stale`;
    expect(resolveSelectionUnit(index, 'body', at('00000001', 7), 'word', stale)).toBeNull();
    expect(resolveSelectionText(index, 'body', at('00000001', 0), at('00000001', 5), stale)).toBeNull();
    expect(
      resolveYrsPointPosition(index, { position: at('00000001', 1), pageIndex: 0, region: 'body' }, stale)
    ).toBeNull();
  });
});
