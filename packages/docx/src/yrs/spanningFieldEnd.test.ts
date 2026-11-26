import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { parseDocx } from '../docx';
import { createYrsSession, saveYrsDocx } from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DOCUMENT_PART = 'word/document.xml';
const BODY_NEXT = 'BODYNEXT';
const INSERTED = 'INSERTED';
const ENTRY_TWO = 'ENTRYTWO';
const PAGE_MARKERS = ['PAGEONE', 'PAGETWO', 'PAGETHREE'] as const;
const S1_EDITS = [
  ['last', 'SPANEND'],
  ['first', 'SPANSTART'],
  ['after', BODY_NEXT],
  ['none', null],
  ['insert_inside', null],
] as const;
const TOC_EDITS = [
  ['last', 'ENTRYTHREE'],
  ['middle', ENTRY_TWO],
  ['first', 'ENTRYONE'],
  ['after', BODY_NEXT],
  ['none', null],
  ['insert_inside', null],
  ['delete_middle', null],
] as const;

const fieldChar = (kind: string) => `<w:r><w:fldChar w:fldCharType="${kind}"/></w:r>`;
const fieldStart = (code: string) =>
  fieldChar('begin') + `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>` + fieldChar('separate');
const textRun = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;

function entry(number: number, marker: string, pageMarker: string): string {
  return `<w:hyperlink w:anchor="_Toc${number}" w:history="1">${textRun(marker)}<w:r><w:tab/></w:r>` +
    `${fieldStart(`PAGEREF _Toc${number} \\h`)}${textRun(pageMarker)}${fieldChar('end')}</w:hyperlink>`;
}

function body(scenario: 'S1' | 'TA' | 'TB'): string {
  const end = fieldChar('end');
  const paragraphs = scenario === 'S1' ? [
    fieldStart('QUOTE') + textRun('SPANSTART'),
    textRun('SPANEND') + end + textRun('AFTERSPAN'),
  ] : [
    fieldStart('TOC \\o "1-3" \\h \\z \\u') + entry(1, 'ENTRYONE', PAGE_MARKERS[0]),
    entry(2, ENTRY_TWO, PAGE_MARKERS[1]),
    entry(3, 'ENTRYTHREE', PAGE_MARKERS[2]) + (scenario === 'TA' ? end : ''),
  ];
  if (scenario === 'TB') paragraphs.push(end);
  paragraphs.push(textRun(BODY_NEXT));
  return paragraphs.map((content) => `<w:p>${content}</w:p>`).join('');
}

function packageBytes(body: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>');
  zip.file('_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>');
  zip.file(DOCUMENT_PART, `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

function decodeXmlText(text: string): string {
  const entities: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  return text.replace(/&(amp|lt|gt|quot|apos|#\d+|#x[\da-fA-F]+);/g, (entity, name: string) => {
    if (!name.startsWith('#')) return entities[name] ?? entity;
    const code = name.startsWith('#x') ? Number.parseInt(name.slice(2), 16) : Number(name.slice(1));
    return code <= 0x10ffff ? String.fromCodePoint(code) : entity;
  });
}

interface FieldSpan {
  begin: number;
  end: number | null;
  separates: number[];
  instructions: string[];
  result: string;
  parent: number | null;
}

function inspect(xml: string) {
  const paragraphs: string[] = [];
  const fields: FieldSpan[] = [];
  const open: FieldSpan[] = [];
  const instructions: string[] = [];
  let ends = 0;
  for (const [, paragraph] of xml.matchAll(/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g)) {
    const index = paragraphs.length;
    let visible = '';
    for (const [, tag, attributes = '', text = ''] of paragraph!.matchAll(/<w:(fldChar|instrText|t)\b([^>]*)>([^<]*)/g)) {
      if (tag === 'fldChar' && attributes.includes('w:fldCharType="begin"')) {
        const parent = open.at(-1);
        const field = { begin: index, end: null, separates: [], instructions: [], result: '',
          parent: parent ? fields.indexOf(parent) : null };
        fields.push(field);
        open.push(field);
      } else if (tag === 'fldChar' && attributes.includes('w:fldCharType="end"')) {
        ends += 1;
        const field = open.pop();
        if (!field) throw new Error('Field end without a begin');
        field.end = index;
      } else if (tag === 'fldChar' && attributes.includes('w:fldCharType="separate"')) {
        const field = open.at(-1);
        if (!field) throw new Error('Field separator without a begin');
        field.separates.push(index);
      } else if (tag === 'instrText') {
        const code = decodeXmlText(attributes.endsWith('/') ? '' : text);
        const field = open.at(-1);
        if (!field) throw new Error('Field instruction without a begin');
        expect(field.separates).toEqual([]);
        field.instructions.push(code);
        instructions.push(code);
      } else if (tag === 't' && !attributes.endsWith('/')) {
        const value = decodeXmlText(text);
        visible += value;
        for (const field of open) {
          if (field.separates.length) field.result += value;
        }
      }
    }
    paragraphs.push(visible);
  }
  expect(open).toEqual([]);
  expect(fields.length).toBe(ends);
  return { paragraphs, fields, instructions };
}

function shiftFields(fields: FieldSpan[], after: number, by: number) {
  const shift = (index: number) => index > after ? index + by : index;
  for (const field of fields) {
    field.begin = shift(field.begin);
    if (field.end !== null) field.end = shift(field.end);
    field.separates = field.separates.map(shift);
  }
}

for (const scenario of ['S1', 'TA', 'TB'] as const) {
  const edits = scenario === 'S1' ? S1_EDITS : TOC_EDITS;
  for (const [edit, marker] of edits) {
    test(`${scenario} preserves field boundaries after session save (${edit})`, async () => {
      const sourceBody = body(scenario);
      const source = await packageBytes(sourceBody);
      const expected = inspect(sourceBody);
      const session = await createYrsSession();
      try {
        session.seedFromDocx(source);
        if (marker !== null) {
          const match = session.searchText(marker, { caseSensitive: true }).find(({ story }) => story === 'body');
          if (match) {
            session.insertText({ story: match.story, paraId: match.paraId, offset: match.end }, ' edited');
          } else if (scenario === 'S1' && edit === 'first') {
            const [paragraph] = session.paragraphSpans('body');
            const [field] = session.storySegments('body');
            if (paragraph?.length !== 1 || field?.kind !== 'embed' || field.embedKind !== 'field' ||
                field.payload.displayText !== marker) throw new Error('Missing cached field result at the edit position');
            session.insertText({ story: 'body', paraId: paragraph.paraId, offset: paragraph.length }, ' edited');
          } else {
            throw new Error('Missing searchable edit marker');
          }
          const replace = (text: string) => text.replace(marker, `${marker} edited`);
          expected.paragraphs = expected.paragraphs.map(replace);
          for (const field of expected.fields) field.result = replace(field.result);
        } else if (edit === 'insert_inside' || edit === 'delete_middle') {
          const story = 'body';
          const index = scenario === 'S1' ? 0 : 1;
          const paragraph = session.paragraphs(story)[index];
          if (!paragraph) throw new Error(`Missing paragraph for ${edit}`);
          const { paraId } = paragraph;
          const result = session.applyEdits({
            expectVersion: session.version(),
            steps: [edit === 'insert_inside'
              ? { op: 'insertParagraphs', target: { story, paraId }, at: 'end', paragraphs: [{ text: INSERTED }] }
              : { op: 'deleteParagraphs', story, firstParaId: paraId, lastParaId: paraId }],
          });
          expect(result.ok).toBe(true);
          const outer = expected.fields[0]!;
          if (edit === 'insert_inside') {
            const offset = expected.paragraphs.slice(0, index + 1).join('').length;
            expected.paragraphs.splice(index + 1, 0, INSERTED);
            outer.result = outer.result.slice(0, offset) + INSERTED + outer.result.slice(offset);
            shiftFields(expected.fields, index, 1);
          } else {
            const [removed] = expected.paragraphs.splice(index, 1);
            const [field] = expected.fields.splice(index + 1, 1);
            outer.result = outer.result.replace(removed!, '');
            expected.instructions = expected.instructions.filter((code) => !field!.instructions.includes(code));
            shiftFields(expected.fields, index, -1);
          }
        }
        const { bytes } = await saveYrsDocx(session, { updateModifiedDate: false });
        const zip = await JSZip.loadAsync(bytes);
        const part = zip.file(DOCUMENT_PART);
        if (!part) throw new Error('Missing word/document.xml');
        const actual = inspect(await part.async('string'));
        expect(actual.paragraphs).toEqual(expected.paragraphs);
        expect(actual.instructions).toEqual(expected.instructions);
        expect(actual.fields).toEqual(expected.fields);
        const reopened = await parseDocx(bytes, { preloadFonts: false });
        const [first] = reopened.package.document.content;
        const field = first?.type === 'paragraph' ? first.content.find(node => node.type === 'complexField') : undefined;
        expect(field?.structuredResult?.blockCount).toBe(expected.fields[0]!.end!);
      } finally {
        session.destroy();
      }
    }, 30_000);
  }
}
