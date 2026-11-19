import { describe, expect, it } from 'bun:test';
import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import { checkedSaveModel, commentRangeWellFormed, type SaveModel } from './commentRangeCheck';

const text = (value: string) => ({ type: 'run', content: [{ type: 'text', text: value }] });
const reference = (id: number) => ({ type: 'run', content: [{ type: 'commentReference', id }] });
const opening = (id: number) => ({ type: 'commentRangeStart', id });
const closing = (id: number) => ({ type: 'commentRangeEnd', id });
const paragraph = (...content: unknown[]) => ({ type: 'paragraph', content });
const cells = (...contents: unknown[][]) => ({
  type: 'table',
  rows: [{ cells: contents.map((content) => ({ content })) }],
});
const box = (xml: string, model?: unknown) => ({
  type: 'run',
  content: [{ type: 'opaqueDrawing', kind: 'drawing', xml, ...(model ? { textBox: model } : {}) }],
});
const boxXml = (inner: string) =>
  `<w:drawing><wp:inline><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent><w:p>${inner}</w:p>` +
  '</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>';
const XML_END = (id: number) => `<w:commentRangeEnd w:id="${id}"/>`;
const XML_REFERENCE = (id: number) => `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
const XML_TEXT = (value: string) => `<w:r><w:t>${value}</w:t></w:r>`;

function model(body: unknown[], extra: { headers?: Map<string, { content: unknown[] }>; comments?: Partial<Comment>[] } = {}) {
  return {
    package: {
      document: { content: body, ...(extra.comments ? { comments: extra.comments } : {}) },
      ...(extra.headers ? { headers: extra.headers } : {}),
    },
  } as unknown as Document;
}

const WHOLE_INTO_A_CELL = model([
  paragraph(text('Alpha '), opening(1), text('one.')),
  cells([paragraph(text('Gamma'), closing(1), reference(1), text(' three.'))]),
]);

describe('whether a comment range in a save model is well-formed', () => {
  it('holds for a range written whole from a body paragraph into a table cell', () => {
    expect(commentRangeWellFormed(WHOLE_INTO_A_CELL, 1)).toBe(true);
  });

  it('holds for a range cut at the end of the paragraph it starts in, its reference right after its end', () => {
    expect(commentRangeWellFormed(model([paragraph(text('Alpha '), opening(1), text('one.'), closing(1), reference(1))]), 1)).toBe(true);
  });

  it('holds for a range cut where it starts whose end the serializer gives its reference', () => {
    expect(commentRangeWellFormed(model([paragraph(text('Alpha '), opening(1), text('one.'), closing(1))]), 1)).toBe(true);
  });

  it('holds for a range into a text box whose markup and typed content both carry its end and reference', () => {
    const inner = `${XML_TEXT('Gamma')}${XML_END(1)}${XML_REFERENCE(1)}`;
    const typed = { textBody: { content: [paragraph(text('Gamma'), closing(1), reference(1))] } };
    expect(commentRangeWellFormed(model([paragraph(opening(1), text('one.')), paragraph(box(boxXml(inner), typed))]), 1)).toBe(true);
  });

  it('fails for a range whose end and reference a dropdown display paragraph replaced', () => {
    const control = { type: 'blockSdt', properties: {}, content: [paragraph(text('Omega'))] };
    expect(commentRangeWellFormed(model([paragraph(text('Alpha '), opening(1), text('one.')), control]), 1)).toBe(false);
  });

  it('fails for a range whose end and reference sit in a table nested in a cell before the cell of its start', () => {
    const merged = [paragraph(text('Top left.')), cells([paragraph(text('Inner'), closing(1), reference(1))]), paragraph()];
    const table = { type: 'table', rows: [{ cells: [{ content: merged }, { content: [paragraph(opening(1), text('Top right.'))] }] }] };
    expect(commentRangeWellFormed(model([table]), 1)).toBe(false);
  });

  it('fails for a range with two ends and its reference in one text box', () => {
    const inner = `${XML_TEXT('Gamma')}${XML_END(1)}${XML_TEXT(' more')}${XML_END(1)}${XML_REFERENCE(1)}`;
    expect(commentRangeWellFormed(model([paragraph(opening(1), text('one.')), paragraph(box(boxXml(inner)))]), 1)).toBe(false);
  });

  it('fails for a range whose text box holds its reference before its end', () => {
    const inner = `${XML_TEXT('Gamma')}${XML_REFERENCE(1)}${XML_END(1)}`;
    expect(commentRangeWellFormed(model([paragraph(opening(1), text('one.')), paragraph(box(boxXml(inner)))]), 1)).toBe(false);
  });

  it('fails for a reply ranged in a header and given a second start in the body', () => {
    const body = [paragraph(opening(1), opening(10), text('one.')), { type: 'blockSdt', properties: {}, content: [paragraph(text('Gamma'), closing(1), reference(1))] }];
    const headers = new Map([['rIdH', { content: [paragraph(text('Head '), opening(10), text('note'), closing(10), reference(10))] }]]);
    const replied = model(body, { headers });
    expect([commentRangeWellFormed(replied, 1), commentRangeWellFormed(replied, 10)]).toEqual([true, false]);
  });

  it('fails for a range whose end is followed by a run holding text and then the reference', () => {
    const shared = { type: 'run', content: [{ type: 'text', text: 'outside' }, { type: 'commentReference', id: 1 }] };
    const saved = model([paragraph(opening(1), text('one.')), cells([paragraph(text('Gamma'), closing(1), shared)])]);
    expect(commentRangeWellFormed(saved, 1)).toBe(false);
  });

  it('fails for a range whose end is followed by text before its reference run', () => {
    const saved = model([paragraph(opening(1), text('one.')), cells([paragraph(text('Gamma'), closing(1), text('outside'), reference(1))])]);
    expect(commentRangeWellFormed(saved, 1)).toBe(false);
  });
});

describe('a save model checked for the ranges it wrote across stories', () => {
  const built = (document: Document, whole: Array<[number, string]>): SaveModel => ({ document, whole: new Map(whole) });

  it('is built once when every range written across stories is well-formed', () => {
    const fallbacks: Array<ReadonlyMap<number, string>> = [];
    const result = checkedSaveModel((fallback) => {
      fallbacks.push(fallback);
      return built(WHOLE_INTO_A_CELL, [[1, 'body']]);
    });
    expect([result.document, fallbacks.map((fallback) => [...fallback])]).toEqual([WHOLE_INTO_A_CELL, [[]]]);
  });

  it('is built once more, the failing comment and the reply on it on the fallback, and kept as it comes', () => {
    const replied = model(
      [paragraph(opening(1), opening(10), text('one.')), cells([paragraph(text('Gamma'), closing(1), reference(1))])],
      {
        headers: new Map([['rIdH', { content: [paragraph(opening(10), text('note'), closing(10), reference(10))] }]]),
        comments: [{ id: 1 }, { id: 10, parentId: 1 }],
      }
    );
    const fallbacks: Array<Array<[number, string]>> = [];
    const result = checkedSaveModel((fallback) => {
      fallbacks.push([...fallback]);
      return built(replied, [[1, 'body'], [10, 'hf:rIdH']]);
    });
    expect([result.document === replied, fallbacks]).toEqual([true, [[], [[1, 'body'], [10, 'hf:rIdH']]]]);
  });
});
