import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';

import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { preloadParseWasm } from '../wasm/parse';
import { applyClipboardFragment, readClipboardSlice } from './clipboardFragment';
import type { DocxEditStep, DocxTextTarget } from './edits';
import { createYrsSession, saveYrsDocx, SuggestedFormattingOverContainerError, type YrsSession, type YrsStoryRange } from './index';
import type { TablePayload } from './storyPlainText';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DATE = '2026-10-01T12:00:00Z';
const AUTHOR = { name: 'Ada', date: DATE };
const sessions: YrsSession[] = [];

function docx(body: string): Uint8Array {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
    )],
    ['_rels/.rels', toBytes(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/>` +
      '</Relationships>'
    )],
    ['word/document.xml', toBytes(
      `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`
    )],
  ])));
}

async function loadWasm(): Promise<void> {
  await Promise.all([
    preloadEditWasm(readFileSync(new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url))),
    preloadOpcWasm(readFileSync(new URL('../wasm/generated/opc/ooxml_opc_bg.wasm', import.meta.url))),
    preloadParseWasm(readFileSync(new URL('../wasm/generated/parse/docx_parse_bg.wasm', import.meta.url))),
  ]);
}

beforeAll(loadWasm, 30_000);
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

function required<T>(value: T | null | undefined): T {
  expect(value).not.toBeNull();
  expect(value).not.toBeUndefined();
  return value!;
}

async function open(body: string): Promise<YrsSession> {
  const session = await createYrsSession();
  sessions.push(session);
  session.openDocx(docx(body), true);
  return session;
}

async function saveXml(session: YrsSession): Promise<string> {
  const { bytes } = await saveYrsDocx(session, { updateModifiedDate: false });
  return new TextDecoder().decode(required(unzipContainer(bytes)['word/document.xml']));
}

function textRange(session: YrsSession, story: string, text: string): YrsStoryRange {
  const match = required(session.searchText(text, { caseSensitive: true }).find(entry => entry.story === story));
  return {
    story,
    start: { paraId: match.paraId, offset: match.start },
    end: { paraId: match.paraId, offset: match.end },
  };
}

function suggestBold(session: YrsSession, range: YrsStoryRange): void {
  session.toggleMark(range, { type: 'bold' }, AUTHOR);
  expect(session.storySegments(range.story).some(segment =>
    segment.kind === 'text' && segment.attributes.rPrChange != null
  )).toBe(true);
}

function fieldParagraph(revision = ''): string {
  return '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> REF entry \\h </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    `<w:hyperlink w:anchor="entry"><w:r><w:rPr><w:i/>${revision}</w:rPr><w:t>Entry</w:t></w:r></w:hyperlink>` +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
}

function savedField(xml: string) {
  const paragraph = required(/<w:p(?=[\s>])[^>]*>[\s\S]*?<\/w:p>/.exec(xml))[0];
  const order = [...paragraph.matchAll(
    /<w:fldChar\b[^>]*\bw:fldCharType="(begin|separate|end)"[^>]*>|<w:hyperlink(?=[\s/>])/g
  )].map(match => match[1] ?? 'hyperlink');
  const result = /<w:fldChar\b[^>]*\bw:fldCharType="separate"[^>]*>([\s\S]*?)<w:fldChar\b[^>]*\bw:fldCharType="end"[^>]*>/
    .exec(paragraph)?.[1] ?? '';
  const text = [...result.matchAll(/<w:t(?=[\s>])[^>]*>([^<]*)<\/w:t>/g)]
    .map(match => match[1]).join('');
  return { order, result, text };
}

const TABLE = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';

function firstCellStory(session: YrsSession, tableIndex: number): string {
  const table = session.tablePayload('body', tableIndex) as TablePayload | null;
  return required(table?.rows[0]?.cells[0]?.story);
}

describe('tracked formatting in fields and copied content', () => {
  it('keeps the REF field result when rejecting an imported run property change', async () => {
    const session = await open(fieldParagraph(
      `<w:rPrChange w:id="41" w:author="Prior" w:date="${DATE}"><w:rPr><w:b/></w:rPr></w:rPrChange>`
    ));
    const before = required(session.storySegments('body').find(segment => segment.kind === 'text'));
    const revision = required(session.listRevisions().find(entry => entry.kind === 'rPrChange' && entry.author === 'Prior'));
    session.rejectChange({ revisionId: revision.revisionId });
    const after = required(session.storySegments('body').find(segment => segment.kind === 'text'));
    expect(after.attributes.fieldResult).toEqual(before.attributes.fieldResult);
    expect(after.attributes.hyperlink).toEqual(before.attributes.hyperlink);

    const xml = await saveXml(session);
    const field = savedField(xml);
    expect(field.order).toEqual(['begin', 'separate', 'hyperlink', 'end']);
    expect(field.text).toBe('Entry');
    expect(field.result).toMatch(/<w:hyperlink(?=[\s>])[^>]*>[\s\S]*?<\/w:hyperlink>/);
    expect(field.result).toMatch(/<w:b(?=[\s/>])/);
    expect(field.result).not.toMatch(/<w:i(?=[\s/>])/);
    expect(xml).not.toContain('w:rPrChange');
    expect(session.listRevisions()).toEqual([]);
  });

  it('refuses a suggested hyperlink over a REF field result without mutation', async () => {
    const session = await open(fieldParagraph());
    const range = textRange(session, 'body', 'Entry');
    const beforeXml = await saveXml(session);
    const before = session.encodeState();
    expect(() => session.setHyperlink(range, { href: '#new-entry' }, AUTHOR))
      .toThrow(SuggestedFormattingOverContainerError);
    expect(session.encodeState()).toEqual(before);
    expect(session.listRevisions()).toEqual([]);
    const xml = await saveXml(session);
    expect(xml).toBe(beforeXml);
    expect(savedField(xml).order).toEqual(['begin', 'separate', 'hyperlink', 'end']);
    expect(savedField(xml).text).toBe('Entry');
    expect(xml).not.toMatch(/<w:(ins|del)(?=[\s/>])/);
  });

  it('refuses host replacement, deletion and paragraph deletion over a live formatting suggestion', async () => {
    for (const op of ['replaceText', 'deleteText', 'deleteParagraphs'] as const) {
      const session = await open('<w:p><w:r><w:t>Hello world</w:t></w:r></w:p><w:p/>');
      const range = textRange(session, 'body', 'Hello');
      suggestBold(session, range);
      const before = session.encodeState();
      const revisions = session.listRevisions();
      const target: DocxTextTarget = {
        kind: 'search', text: 'Hello', within: { kind: 'story', story: 'body' }, view: 'accepted',
      };
      const step: DocxEditStep = op === 'deleteParagraphs'
        ? { op, story: 'body', firstParaId: range.start.paraId, lastParaId: range.start.paraId }
        : op === 'replaceText' ? { op, target, text: 'Hi' } : { op, target };
      const result = session.applyEdits({
        expectVersion: session.version(),
        source: 'host',
        steps: [step],
      });

      expect(result.ok).toBe(false);
      if (result.ok) throw new Error(`${op} was applied`);
      expect(result.failure.code).toBe('tracked-revision-conflict');
      expect(session.encodeState()).toEqual(before);
      expect(session.listRevisions()).toEqual(revisions);
      expect(session.listRevisions().some(revision => revision.kind === 'rPrChange')).toBe(true);
      expect(await saveXml(session)).toContain('w:rPrChange');
    }
  });

  it('gives copied table formatting suggestions independent revision ids', async () => {
    const session = await open(`${TABLE}<w:p/>`);
    const originalStory = firstCellStory(session, 0);
    suggestBold(session, textRange(session, originalStory, 'Hello'));
    const originalRevision = required(session.listRevisions().find(revision =>
      revision.story === originalStory && revision.kind === 'rPrChange'
    ));
    const anchor = required(session.paragraphs('body').findLast(paragraph => paragraph.text === ''));
    const span = session.locateParagraph('body', anchor.paraId);
    const at = { paraId: anchor.paraId, offset: span.end - span.start };
    const cell = { story: 'body', tableIndex: 0, row: 0, column: 0 };
    const { fragment } = readClipboardSlice(session, { cells: { anchor: cell, head: cell } });
    applyClipboardFragment(session, { story: 'body', start: at, end: at }, required(fragment));
    const pastedStory = required([firstCellStory(session, 0), firstCellStory(session, 1)]
      .find(story => story !== originalStory));
    const pastedRevision = required(session.listRevisions().find(revision =>
      revision.story === pastedStory && revision.kind === 'rPrChange'
    ));

    expect(pastedRevision.revisionId).not.toBe(originalRevision.revisionId);
    session.rejectChange({ revisionId: pastedRevision.revisionId });
    expect(session.selectionContext(textRange(session, originalStory, 'Hello')).bold).toBe(true);
    expect(session.selectionContext(textRange(session, pastedStory, 'Hello')).bold).toBe(false);
    expect(session.listRevisions()).toContainEqual(originalRevision);
    expect(session.listRevisions().some(revision => revision.revisionId === pastedRevision.revisionId)).toBe(false);
    expect(await saveXml(session)).toContain('w:rPrChange');
  });
});
