import { test } from 'bun:test';
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

function outerEndParagraph(xml: string): number | 'none' {
  const fields: boolean[] = [];
  for (const [index, paragraph] of xml.split('</w:p>').entries()) {
    for (const [, tag, attributes = '', text = ''] of paragraph.matchAll(/<w:(fldChar|instrText)\b([^>]*)>([^<]*)/g)) {
      if (tag === 'fldChar') {
        if (attributes.includes('w:fldCharType="begin"')) fields.push(false);
        else if (attributes.includes('w:fldCharType="end"') && fields.pop()) return index;
      } else if (fields.length > 0 && /^(?:QUOTE|TOC)\b/.test(text.trimStart())) {
        fields[fields.length - 1] = true;
      }
    }
  }
  return 'none';
}

function hasMarker(value: unknown, marker: string): boolean {
  if (Array.isArray(value)) return value.some((child) => hasMarker(child, marker));
  if (value === null || typeof value !== 'object') return false;
  const node = value as Record<string, unknown>;
  return (node.type === 'text' && typeof node.text === 'string' && node.text.includes(marker)) ||
    Object.values(node).some((child) => hasMarker(child, marker));
}

function fieldContainsBodyNext(value: unknown, following: readonly unknown[]): boolean {
  if (Array.isArray(value)) return value.some((child) => fieldContainsBodyNext(child, following));
  if (value === null || typeof value !== 'object') return false;
  const node = value as Record<string, unknown>;
  if (node.type === 'complexField') {
    const code = node.structuredCode as { blockCount?: number } | undefined;
    const result = node.structuredResult as { blockCount?: number } | undefined;
    const start = code?.blockCount ?? 0;
    const end = start + (result?.blockCount ?? 0);
    if (hasMarker(result, BODY_NEXT) || hasMarker(node.fieldResult, BODY_NEXT) ||
        following.slice(start, end).some((block) => hasMarker(block, BODY_NEXT))) return true;
  }
  return Object.values(node).some((child) => fieldContainsBodyNext(child, following));
}

function nestedFlags(present: (marker: string) => boolean): string {
  return PAGE_MARKERS.map((marker) => present(marker) ? 'y' : 'n').join('');
}

test('measure paragraph-spanning field ends after a session save', async () => {
  for (const scenario of ['S1', 'TA', 'TB'] as const) {
    const edits = scenario === 'S1' ? S1_EDITS : TOC_EDITS;
    for (const [edit, marker] of edits) {
      let begins = 0;
      let ends = 0;
      let endPara: number | 'none' = 'none';
      let textOk = false;
      let reopen: 'swallowed' | 'ok' | 'err' = 'err';
      let nested = 'nnn';
      let nestedModel = 'nnn';
      let routeNote = 'n/a';
      try {
        const source = await packageBytes(body(scenario));
        const session = await createYrsSession();
        try {
          session.seedFromDocx(source);
          if (marker !== null) {
            const match = session.searchText(marker, { caseSensitive: true }).find(({ story }) => story === 'body');
            if (!match) throw new Error(`Missing edit marker ${marker}`);
            session.insertText({ story: match.story, paraId: match.paraId, offset: match.end }, ' edited');
          } else if (edit === 'insert_inside' || edit === 'delete_middle') {
            const story = 'body';
            const paragraph = session.paragraphs(story)[scenario === 'S1' ? 0 : 1];
            if (!paragraph) throw new Error(`Missing paragraph for ${edit}`);
            const { paraId } = paragraph;
            const result = session.applyEdits({
              expectVersion: session.version(),
              steps: [edit === 'insert_inside'
                ? { op: 'insertParagraphs', target: { story, paraId }, at: 'end', paragraphs: [{ text: INSERTED }] }
                : { op: 'deleteParagraphs', story, firstParaId: paraId, lastParaId: paraId }],
            });
            if (!result.ok) throw new Error(`Session refused ${edit}`);
          }
          const { bytes, structureDropped } = await saveYrsDocx(session, { updateModifiedDate: false });
          if (structureDropped?.length) routeNote = `whole:${structureDropped.map(({ part }) => part).join(',')}`;
          const zip = await JSZip.loadAsync(bytes);
          const part = zip.file(DOCUMENT_PART);
          if (!part) throw new Error('Missing word/document.xml');
          const xml = await part.async('string');
          begins = (xml.match(/w:fldCharType="begin"/g) ?? []).length;
          ends = (xml.match(/w:fldCharType="end"/g) ?? []).length;
          endPara = outerEndParagraph(xml);
          const text = Array.from(xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g), (match) => match[1]).join('');
          nested = nestedFlags((target) => text.includes(target));
          const deletedMiddle = edit === 'delete_middle';
          const expectedNested = deletedMiddle ? 'yny' : 'yyy';
          const markers = edits.flatMap(([, target]) =>
            target === null || (deletedMiddle && target === ENTRY_TWO) ? [] : [target]);
          textOk = markers.every((target) => text.includes(target === marker ? `${target} edited` : target)) &&
            (scenario === 'S1' ? text.includes('AFTERSPAN') : nested === expectedNested) &&
            (edit !== 'insert_inside' || text.includes(INSERTED)) &&
            (!deletedMiddle || !text.includes(ENTRY_TWO));
          const document = await parseDocx(bytes, { preloadFonts: false });
          const blocks = document.package.document.content;
          const swallowed = blocks.some((block, index) => fieldContainsBodyNext(block, blocks.slice(index + 1)));
          reopen = swallowed ? 'swallowed' : 'ok';
          nestedModel = nestedFlags((target) => blocks.some((block) => hasMarker(block, target)));
        } finally {
          session.destroy();
        }
      } catch {
        reopen = 'err';
      }
      console.log(`SPAN session ${scenario} ${edit} | begins=${begins} ends=${ends} | end_para=${endPara} | ` +
        `text_ok=${textOk ? 'yes' : 'no'} | reopen=${reopen} | nested=${nested} | nested_model=${nestedModel} | route_note=${routeNote}`);
    }
  }
}, 30_000);
