import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { DisplayPositionIndex } from './displayPositionIndex';
import { createYrsSession, type YrsSession } from './index';
import { resolveYrsPointPosition } from './pointPosition';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';
import { createProposalRegistry } from './proposals';
import {
  resolveBookmarkPosition,
  resolveCommentTarget,
  resolveParagraphTarget,
  resolveRevisionTarget,
  resolveSelectionInfo,
  resolveSelectionText,
  resolveSelectionUnit,
} from './viewerSelection';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const p = (id: string, content: string) => `<w:p w14:paraId="${id}">${content}</w:p>`;
const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const cell = (content: string) => `<w:tc><w:tcPr/>${content}</w:tc>`;

function docx(body: string, comments?: string): Uint8Array {
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr/></w:body></w:document>`)],
  ]);
  if (comments !== undefined) {
    const types = new TextDecoder().decode(parts.get('[Content_Types].xml')!);
    parts.set('[Content_Types].xml', toBytes(types.replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')));
    parts.set('word/_rels/document.xml.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>'));
    parts.set('word/comments.xml', toBytes(`<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${comments}</w:comments>`));
  }
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
let main: YrsSession;
let index: DisplayPositionIndex;

beforeAll(async () => {
  preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  resident = await createResidentEngineSession();
  resident.openDocx(docx(BODY));
  main = await createYrsSession();
  main.loadState(resident.encodeState());
  index = new DisplayPositionIndex({
    ...resident.geometryReader,
    selectionText: resident.selectionText,
    resolveComment: resident.resolveComment,
  });
});

afterAll(() => {
  main.destroy();
  resident.destroy();
});

async function withDocument(body: string, check: (resident: ResidentEngineSession, index: DisplayPositionIndex) => void, comments?: string) {
  const resident = await createResidentEngineSession();
  try {
    resident.openDocx(docx(body, comments));
    const index = new DisplayPositionIndex({
      ...resident.geometryReader,
      selectionText: resident.selectionText,
      resolveComment: resident.resolveComment,
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

  test('selection info matches the main session for single, multiple and collapsed paragraphs', () => {
    const version = resident.geometryReader.version();
    const ranges = [
      { start: { paraId: '00000001', offset: 6 }, end: { paraId: '00000001', offset: 10 } },
      { start: { paraId: '00000001', offset: 6 }, end: { paraId: '00000002', offset: 13 } },
      { start: { paraId: '00000001', offset: 6 }, end: { paraId: '00000001', offset: 6 } },
    ];
    for (const range of ranges) {
      const anchor = at(range.start.paraId, range.start.offset);
      const head = at(range.end.paraId, range.end.offset);
      const expected = main.selectionText({ story: 'body', ...range });
      expect(resolveSelectionInfo(index, 'body', anchor, head, version)).toEqual(expected);
      expect(resolveSelectionInfo(index, 'body', head, anchor, version)).toEqual(expected);
      expect(resolveSelectionInfo(index, 'body', anchor, head, `${version}-stale`)).toBeNull();
    }
  });

  test('selection info preserves the visible text across two table cells', () => {
    const version = resident.geometryReader.version();
    const anchor = at('00000003', 1);
    const head = at('00000004', 1);
    expect(resolveSelectionText(index, 'body', anchor, head, version)?.text).toBe('A1\tB1');
    const expected = { paraId: null, selectedText: 'A1\tB1', paragraphText: '', before: '', after: '' };
    expect(resolveSelectionInfo(index, 'body', anchor, head, version)).toEqual(expected);
    expect(resolveSelectionInfo(index, 'body', head, anchor, version)).toEqual(expected);
    expect(resolveSelectionInfo(index, 'body', anchor, head, `${version}-stale`)).toBeNull();
  });

  test('selection info preserves cross-story text and rejects unresolved display ranges', () => {
    const version = resident.geometryReader.version();
    const anchor = at('00000001', 1);
    const head = at('00000003', 1);
    expect(resolveSelectionInfo(index, 'body', anchor, head, version)).toEqual({
      paraId: null, selectedText: resolveSelectionText(index, 'body', anchor, head, version)!.text,
      paragraphText: '', before: '', after: '',
    });
    expect(resolveSelectionInfo(index, 'missing', 0, 1, version)).toBeNull();
    expect(resolveSelectionInfo(index, 'body', -1, at('00000001', 1), version)).toBeNull();
    expect(resolveSelectionInfo(index, 'body', 0, index.projection('body')!.size + 1, version)).toBeNull();
  });

  test('paragraph targets include table cells and reject unknown, outside and stale targets', () => {
    const version = resident.geometryReader.version();
    for (const paraId of ['00000001', '00000003', '00000006']) {
      const story = resident.geometryReader.storyIds().find((id) => resident.geometryReader.paragraphs(id).some((paragraph) => paragraph.paraId === paraId))!;
      const span = resident.geometryReader.locateParagraph(story, paraId);
      const anchor = index.positionOf({ story, paraId, offset: 0 }, 'body');
      const head = index.positionOf({ story, paraId, offset: span.end - span.start }, 'body');
      expect(anchor).not.toBeNull();
      expect(head).not.toBeNull();
      expect(resolveParagraphTarget(index, 'body', paraId, version)).toEqual({ anchor: anchor!, head: head! });
      expect(resolveParagraphTarget(index, 'missing', paraId, version)).toBeNull();
      expect(resolveParagraphTarget(index, 'body', paraId, `${version}-stale`)).toBeNull();
    }
    expect(resolveParagraphTarget(index, 'body', 'missing', version)).toBeNull();
    expect(resolveParagraphTarget(index, cellStory('00000003'), '00000001', version)).toBeNull();
  });

  test('comment targets map the anchored phrase and reject unknown, outside and stale targets', async () => {
    const body = p('00000001', r('Before ') + '<w:commentRangeStart w:id="1"/>' +
      r('the phrase') + '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>' + r(' after')) +
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell(p('00000003', r('Cell')))}</w:tr></w:tbl>`;
    const comments = '<w:comment w:id="1" w:author="Reviewer"><w:p><w:r><w:t>Check phrase</w:t></w:r></w:p></w:comment>';
    await withDocument(body, (resident, index) => {
      const version = resident.geometryReader.version();
      const position = (offset: number) => index.positionOf({ story: 'body', paraId: '00000001', offset }, 'body')!;
      expect(resolveCommentTarget(index, 'body', '1', version)).toEqual({ anchor: position(7), head: position(17) });
      expect(resolveCommentTarget(index, 'body', 'unknown', version)).toBeNull();
      const otherStory = resident.geometryReader.storyIds().find((story) => story.startsWith('body:'))!;
      expect(resolveCommentTarget(index, otherStory, '1', version)).toBeNull();
      expect(resolveCommentTarget(index, 'missing', '1', version)).toBeNull();
      expect(resolveCommentTarget(index, 'body', '1', `${version}-stale`)).toBeNull();
    }, comments);
  });

  test('revision targets map tracked insertions and reject unknown, outside and stale targets', async () => {
    const body = p('00000001', r('Before ') +
      `<w:ins w:id="9" w:author="A" w:date="2026-10-01T00:00:00Z">${r('new')}</w:ins>` + r(' after')) +
      `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell(p('00000003', r('Cell')))}</w:tr></w:tbl>`;
    await withDocument(body, (resident, index) => {
      const version = resident.geometryReader.version();
      const revision = resident.geometryReader.listRevisions().find((candidate) => candidate.kind === 'insertion')!;
      const position = (offset: number) => index.positionOf({ story: 'body', paraId: '00000001', offset }, 'body')!;
      expect(resolveRevisionTarget(index, 'body', revision.revisionId, version)).toEqual({ anchor: position(7), head: position(10) });
      expect(resolveRevisionTarget(index, 'body', 'unknown', version)).toBeNull();
      const otherStory = resident.geometryReader.storyIds().find((story) => story.startsWith('body:'))!;
      expect(resolveRevisionTarget(index, otherStory, revision.revisionId, version)).toBeNull();
      expect(resolveRevisionTarget(index, 'missing', revision.revisionId, version)).toBeNull();
      expect(resolveRevisionTarget(index, 'body', revision.revisionId, `${version}-stale`)).toBeNull();
    });
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
