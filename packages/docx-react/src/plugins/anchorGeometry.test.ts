import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { createYrsSession, type DocxTextRange, type YrsStorySegment } from '@betteroffice/docx/yrs';
import { textRangeToRaw } from './anchorGeometry';

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
});
