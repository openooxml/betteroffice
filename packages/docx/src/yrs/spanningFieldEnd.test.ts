import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, saveYrsDocx, type YrsAuthor, type YrsRawOp, type YrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

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
const paragraph = (content: string) => `<w:p>${content}</w:p>`;
const blockSdt = (content: string) =>
  '<w:sdt><w:sdtPr><w:id w:val="42"/><w:docPartObj>' +
  '<w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/>' +
  `</w:docPartObj></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

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
  return paragraphs.map(paragraph).join('');
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
  zip.file(DOCUMENT_PART, `<w:document xmlns:w="${W}" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array' });
}

async function readDocument(bytes: Uint8Array | ArrayBuffer): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const part = zip.file(DOCUMENT_PART);
  if (!part) throw new Error('Missing word/document.xml');
  return part.async('string');
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

function storyParagraphs(xml: string) {
  return xml.matchAll(/<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g);
}

function attribute(tag: string, name: string): string | undefined {
  return tag.split(` ${name}="`)[1]?.split('"')[0];
}

function scanStory(xml: string) {
  const paragraphs: string[] = [];
  const fields: FieldSpan[] = [];
  const open: number[] = [];
  const instructions: string[] = [];
  const tabs: number[] = [];
  const endsAt: Array<number | null> = [];
  const markers: string[] = [];
  const markerFields: number[][] = [];
  const bookmarkNames = new Map<string, string>();
  const revisions: string[] = [];
  const paragraphStack: number[] = [];
  let visible = '';
  let runDepth = 0;
  let ends = 0;
  let unpaired = 0;
  for (const [, opening = '', content = ''] of xml.matchAll(/<([^>]*)>([^<]*)/g)) {
    const tag = opening.replace(/\/$/, '').split(/\s+/, 1)[0];
    if (tag === 'w:p') {
      if (!opening.endsWith('/')) paragraphStack.push(paragraphs.length);
      paragraphs.push('');
      tabs.push(0);
    } else if (tag === '/w:p') {
      paragraphStack.pop();
    } else if (tag === 'w:r' && !opening.endsWith('/')) {
      runDepth += 1;
    } else if (tag === '/w:r') {
      runDepth -= 1;
    }

    let id = attribute(opening, 'w:id') ?? '-';
    let kind: string | undefined;
    switch (tag) {
      case 'w:bookmarkStart':
        kind = 'bs';
        bookmarkNames.set(id, attribute(opening, 'w:name') ?? id);
        id = bookmarkNames.get(id)!;
        break;
      case 'w:bookmarkEnd':
        kind = 'be';
        id = bookmarkNames.get(id) ?? id;
        break;
      case 'w:commentRangeStart': kind = 'cs'; break;
      case 'w:commentRangeEnd': kind = 'ce'; break;
      case 'w:commentReference': kind = 'cr'; break;
      case 'w:moveFromRangeStart':
      case 'w:moveFromRangeEnd':
      case 'w:moveToRangeStart':
      case 'w:moveToRangeEnd':
        kind = tag.slice(2);
        break;
      case 'w:ins':
      case 'w:del':
      case 'w:moveFrom':
      case 'w:moveTo':
        kind = tag.slice(2);
        if (!opening.endsWith('/')) revisions.push(id);
        break;
      case '/w:ins':
      case '/w:del':
      case '/w:moveFrom':
      case '/w:moveTo':
        kind = '/' + tag.slice(3);
        id = revisions.pop() ?? '-';
        break;
    }
    if (kind !== undefined) {
      markers.push(`${kind}:${id}@${visible.length}`);
      markerFields.push(open.slice());
    }

    const index = paragraphStack.at(-1);
    if (index === undefined) continue;
    const showsResult = open.every(field => fields[field]!.separates.length > 0);
    if (tag === 'w:fldChar' && opening.includes('w:fldCharType="begin"')) {
      fields.push({ begin: index, end: null, separates: [], instructions: [], result: '',
        parent: open.at(-1) ?? null });
      endsAt.push(null);
      open.push(fields.length - 1);
    } else if (tag === 'w:fldChar' && opening.includes('w:fldCharType="end"')) {
      ends += 1;
      const field = open.pop();
      if (field === undefined) unpaired += 1;
      else {
        fields[field]!.end = index;
        endsAt[field] = visible.length;
      }
    } else if (tag === 'w:fldChar' && opening.includes('w:fldCharType="separate"')) {
      const field = open.at(-1);
      if (field === undefined) unpaired += 1;
      else fields[field]!.separates.push(index);
    } else if (tag === 'w:instrText' && !opening.endsWith('/')) {
      const code = decodeXmlText(content);
      const field = open.at(-1);
      if (field === undefined) unpaired += 1;
      else {
        if (fields[field]!.separates.length) unpaired += 1;
        fields[field]!.instructions.push(code);
      }
      instructions.push(code);
    } else if (tag === 'w:t' && showsResult && !opening.endsWith('/')) {
      const value = decodeXmlText(content);
      visible += value;
      paragraphs[index] += value;
      for (const field of open) {
        if (fields[field]!.separates.length) fields[field]!.result += value;
      }
    } else if (tag === 'w:tab' && showsResult && runDepth > 0) {
      tabs[index] = tabs[index]! + 1;
      visible += '\t';
    }
  }
  return { story: { paragraphs, fields, instructions }, ends, endsAt, visible, tabs, markers,
    markerFields, unpaired: unpaired + open.length };
}

function inspect(xml: string) {
  const saved = scanStory(xml);
  expect(saved.unpaired).toBe(0);
  return saved.story;
}

function assertSpanningSave(xml: string, expectedXml: string, caseName: string): void {
  const actual = scanStory(xml);
  const expected = scanStory(expectedXml);
  const pageRefs = (fields: FieldSpan[]) => fields.flatMap(field => {
    const code = field.instructions.join('');
    return field.end !== null && code.trim().split(/\s+/, 1)[0] === 'PAGEREF' ? [code] : [];
  });
  const actualRefs = pageRefs(actual.story.fields);
  const expectedRefs = pageRefs(expected.story.fields);
  let kept = 0;
  for (const code of expectedRefs) {
    const index = actualRefs.indexOf(code);
    if (index < 0) continue;
    actualRefs.splice(index, 1);
    kept += 1;
  }
  const instr = JSON.stringify(actual.story.instructions) === JSON.stringify(expected.story.instructions)
    ? 'same' : 'changed';
  const endPara = actual.story.fields[0]?.end ?? 'none';
  const expectedEnd = expected.endsAt[0];
  const after = expectedEnd == null ? -1 : actual.visible.lastIndexOf(expected.visible.slice(expectedEnd));
  const end = actual.endsAt[0];
  const outside = end != null && after >= 0 && end <= after;
  console.log(
    `SPANDIFF session ${caseName} begins=${actual.story.fields.length} ends=${actual.ends} instr=${instr} ` +
    `pageref=${kept}/${expectedRefs.length} endpara=${endPara} after=${outside ? 'outside' : 'inside'} ` +
    `markers=${actual.markers.join(',') || '-'}`,
  );
  expect(actual.unpaired).toBe(expected.unpaired);
  expect(actual.story).toEqual(expected.story);
  expect(actual.tabs).toEqual(expected.tabs);
  expect(actual.endsAt).toEqual(expected.endsAt);
  expect(actual.markers).toEqual(expected.markers);
  expect(actual.markerFields).toEqual(expected.markerFields);
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
        const actual = inspect(await readDocument(bytes));
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

function insertInText(session: YrsSession, marker: string, offset: number, text: string, storyId = 'body'): void {
  const match = session.searchText(marker, { caseSensitive: true }).find(({ story }) => story === storyId);
  if (!match) throw new Error(`Missing editable text ${marker}`);
  session.insertText({ story: match.story, paraId: match.paraId, offset: match.start + offset }, text);
}

for (const fullWrite of [false, true]) {
  test(`a block SDT keeps its TOC on ${fullWrite ? 'whole' : 'session'} saves`, async () => {
    const source = blockSdt(body('TA').replace(paragraph(textRun(BODY_NEXT)), '')) +
      paragraph(textRun(BODY_NEXT));
    const story = 'body:sdt0';
    for (const marker of ['ENTRYONE', 'ENTRYTHREE']) {
      const xml = await editAndSave(source, session => {
        insertInText(session, marker, 1, ' EDIT ', story);
      }, fullWrite);
      expect(inspect(xml)).toEqual(inspect(source.replace(marker, marker[0] + ' EDIT ' + marker.slice(1))));
      expect(inspect(xml.split('</w:sdtContent>')[0]!).fields[0]?.end).toBe(2);
    }
    const middle = paragraph(entry(2, ENTRY_TWO, PAGE_MARKERS[1]));
    for (const edit of ['insert', 'delete'] as const) {
      const xml = await editAndSave(source, session => {
        const { paraId } = session.paragraphs(story)[1]!;
        const result = session.applyEdits({
          expectVersion: session.version(),
          steps: [edit === 'insert'
            ? { op: 'insertParagraphs', target: { story, paraId }, at: 'end', paragraphs: [{ text: INSERTED }] }
            : { op: 'deleteParagraphs', story, firstParaId: paraId, lastParaId: paraId }],
        });
        expect(result.ok).toBe(true);
      }, fullWrite);
      const expected = source.replace(middle, edit === 'insert' ? middle + paragraph(textRun(INSERTED)) : '');
      expect(inspect(xml)).toEqual(inspect(expected));
      expect(inspect(xml.split('</w:sdtContent>')[0]!).fields[0]?.end).toBe(edit === 'insert' ? 3 : 1);
    }
  }, 30_000);
}

test('deleting the sole continuation puts the carried end after an empty field result', async () => {
  const opening = textRun('BEFORE') + fieldStart('QUOTE');
  const tail = paragraph(textRun(BODY_NEXT));
  const source = paragraph(opening) + paragraph(fieldChar('end')) + tail;
  for (const fullWrite of [false, true]) {
    const xml = await editAndSave(source, session => {
      const { paraId } = session.paragraphs('body')[1]!;
      const result = session.applyEdits({
        expectVersion: session.version(),
        steps: [{ op: 'deleteParagraphs', story: 'body', firstParaId: paraId, lastParaId: paraId }],
      });
      expect(result.ok).toBe(true);
    }, fullWrite);
    expect(inspect(xml)).toEqual(inspect(paragraph(opening + fieldChar('end')) + tail));
  }
}, 30_000);

async function editAndSave(
  source: string,
  edit: (session: YrsSession) => void,
  fullWrite = false,
  seeder: 'native' | 'projected' = 'native',
): Promise<string> {
  const bytes = await packageBytes(source);
  const session = await createYrsSession();
  try {
    if (seeder === 'native') {
      session.seedFromDocx(bytes);
    } else {
      session.openDocx(bytes, false);
      documentToYrs(session, await parseDocx(bytes, { preloadFonts: false }));
    }
    edit(session);
    if (fullWrite) {
      const base = session.materializeDocx();
      if (!base) throw new Error('Missing session source model');
      return readDocument(await repackDocx(yrsToDocument(session, base), { updateModifiedDate: false }));
    }
    return readDocument((await saveYrsDocx(session, { updateModifiedDate: false })).bytes);
  } finally {
    session.destroy();
  }
}

const WORD_TOC_CODE = 'TOC \\o "1-3" \\h \\z \\u';
const WORD_TOC_TITLES = ['ENTRYONE', ENTRY_TWO, 'ENTRYTHREE'];
const WORD_FIGURE_TITLES = ['FIGUREONE', 'FIGURETWO', 'FIGURETHREE'];

interface WordToc {
  paragraphs: string[];
  titles: readonly string[];
  story: 'body' | 'body:sdt0';
}

interface WordTocOptions {
  code?: string;
  titles?: readonly string[];
  firstAnchor?: number;
  inlineEnd?: boolean;
  numbered?: boolean;
  unlinked?: boolean;
  omitLevelOnePages?: boolean;
}

function wordToc({
  code = WORD_TOC_CODE,
  titles = WORD_TOC_TITLES,
  firstAnchor = 1,
  inlineEnd = false,
  numbered = false,
  unlinked = false,
  omitLevelOnePages = false,
}: WordTocOptions = {}): WordToc {
  const tab = '<w:r><w:rPr><w:webHidden/></w:rPr><w:tab/></w:r>';
  const page = '<w:r><w:rPr><w:webHidden/></w:rPr><w:t>3</w:t></w:r>';
  const end = fieldChar('end');
  const paragraphs = titles.map((title, index) => {
    const level = numbered || omitLevelOnePages ? [1, 2, 1][index] : index + 1;
    const anchor = '_Toc' + (firstAnchor + index);
    let content = `<w:pPr><w:pStyle w:val="TOC${level}"/></w:pPr>`;
    if (index === 0) content += fieldStart(code);
    if (!unlinked) content += `<w:hyperlink w:anchor="${anchor}" w:history="1">`;
    if (numbered) content += textRun(['1', '1.1', '2'][index]!) + tab;
    content += textRun(title);
    if (!omitLevelOnePages || level !== 1) {
      content += tab + fieldStart(`PAGEREF ${anchor} \\h`) + page + end;
    }
    if (!unlinked) content += '</w:hyperlink>';
    if (inlineEnd && index === titles.length - 1) content += end;
    return content;
  });
  if (!inlineEnd) paragraphs.push(end);
  return { paragraphs, titles, story: 'body' };
}

const WORD_TOC_SHAPES = [
  ['W1 hyperlinked plain', () => wordToc()],
  ['W2 inline end', () => wordToc({ inlineEnd: true })],
  ['W3 numbered', () => wordToc({ numbered: true })],
  ['W4 unlinked', () => wordToc({ code: 'TOC \\o "1-3"', unlinked: true })],
  ['W5 without level-one pages', () => wordToc({
    code: WORD_TOC_CODE + ' \\n "1-1"', omitLevelOnePages: true,
  })],
  ['W6 block content control', (): WordToc => {
    const toc = wordToc();
    const heading = '<w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>' + textRun('Contents');
    return { ...toc, paragraphs: [heading, ...toc.paragraphs], story: 'body:sdt0' };
  }],
  ['W7 followed by table of figures', (): WordToc => {
    const toc = wordToc();
    const figures = wordToc({ code: 'TOC \\h \\z \\c "Figure"', titles: WORD_FIGURE_TITLES, firstAnchor: 4 });
    return {
      ...toc,
      paragraphs: [...toc.paragraphs, ...figures.paragraphs],
      titles: [...toc.titles, ...figures.titles],
    };
  }],
  ['W8 single entry', () => wordToc({ titles: WORD_TOC_TITLES.slice(0, 1) })],
] as const;

type WordTocEdit =
  | { kind: 'text'; story: WordToc['story']; marker: string }
  | { kind: 'insert' | 'delete'; story: WordToc['story']; index: number };

const WORD_TOC_EDITS: ReadonlyArray<readonly [string, (toc: WordToc) => WordTocEdit | null]> = [
  ['E1 edit first title', toc => ({ kind: 'text', story: toc.story, marker: toc.titles[0]! })],
  ['E2 edit last title', toc => ({ kind: 'text', story: toc.story, marker: toc.titles.at(-1)! })],
  ['E3 edit middle title', toc => ({ kind: 'text', story: toc.story, marker: toc.titles[1] ?? toc.titles[0]! })],
  ['E4 insert after first entry', toc => ({
    kind: 'insert', story: toc.story, index: toc.story === 'body' ? 0 : 1,
  })],
  ['E5 delete middle entry', toc => toc.titles.length === 1 ? null : {
    kind: 'delete', story: toc.story, index: toc.story === 'body' ? 1 : 2,
  }],
  ['E6 edit following body paragraph', () => ({ kind: 'text', story: 'body', marker: BODY_NEXT })],
];

function wordTocXml(toc: WordToc, paragraphs = toc.paragraphs): string {
  const entries = paragraphs.map(paragraph).join('');
  return (toc.story === 'body' ? entries : blockSdt(entries)) + paragraph(textRun(BODY_NEXT));
}

function applyWordTocEdit(session: YrsSession, edit: WordTocEdit): boolean {
  const { story } = edit;
  if (edit.kind === 'text') {
    const match = session.searchText(edit.marker, { caseSensitive: true }).find(hit => hit.story === story);
    if (!match) return false;
    const receipt = session.insertText({ story, paraId: match.paraId, offset: match.start + 1 }, ' EDIT ');
    return receipt.range !== undefined;
  }
  const target = session.paragraphs(story)[edit.index];
  if (!target) return false;
  const { paraId } = target;
  return session.applyEdits({
    expectVersion: session.version(),
    steps: [edit.kind === 'insert'
      ? { op: 'insertParagraphs', target: { story, paraId }, at: 'end', paragraphs: [{ text: INSERTED }] }
      : { op: 'deleteParagraphs', story, firstParaId: paraId, lastParaId: paraId }],
  }).ok;
}

function editedWordTocXml(toc: WordToc, edit: WordTocEdit, source: string): string {
  if (edit.kind === 'text') {
    const { marker } = edit;
    return source.replace(`>${marker}<`, `>${marker[0]} EDIT ${marker.slice(1)}<`);
  }
  const paragraphs = toc.paragraphs.slice();
  paragraphs.splice(edit.index + (edit.kind === 'insert' ? 1 : 0), edit.kind === 'delete' ? 1 : 0,
    ...(edit.kind === 'insert' ? [textRun(INSERTED)] : []));
  return wordTocXml(toc, paragraphs);
}

async function assertWordTocSave(
  toc: WordToc,
  edit: WordTocEdit,
  caseName: string,
  seeder: 'native' | 'projected' = 'native',
) {
  const source = wordTocXml(toc);
  const expected = editedWordTocXml(toc, edit, source);
  let applied = false;
  const xml = await editAndSave(source, session => {
    applied = applyWordTocEdit(session, edit);
  }, false, seeder);
  assertSpanningSave(xml, expected, caseName);
  expect(applied).toBe(true);
  if (toc.story === 'body:sdt0') {
    const paragraphsInControl = (body: string) => {
      const content = body.match(/<w:sdtContent>([\s\S]*?)<\/w:sdtContent>/)?.[1];
      return content === undefined ? null : [...storyParagraphs(content)].length;
    };
    expect(paragraphsInControl(xml)).toBe(paragraphsInControl(expected));
  }
}

const WORD_TOC_CASES = WORD_TOC_SHAPES.flatMap(([shape, build]) => {
  const toc = build();
  return WORD_TOC_EDITS.flatMap(([name, plan]) => {
    const edit = plan(toc);
    return edit ? [[shape, name, toc, edit] as const] : [];
  });
});

test.each(WORD_TOC_CASES)('%s preserves field boundaries after session save (%s)', async (shape, name, toc, edit) => {
  await assertWordTocSave(toc, edit, `${shape.split(' ')[0]} ${name}`);
}, 30_000);

interface MainState {
  shape: string;
  docx: string;
  update: string;
}

async function mainState(shape: string): Promise<MainState> {
  const fixture = Bun.file(new URL('./__fixtures__/spanning-main-states.json', import.meta.url));
  if (!await fixture.exists()) {
    throw new Error('Missing __fixtures__/spanning-main-states.json. Generate it on main (504f163ef) with .gen-main-states.test.ts.');
  }
  const states = await fixture.json() as MainState[];
  const state = states.find(entry => entry.shape === shape);
  if (!state) throw new Error(`Missing ${shape} in spanning-main-states.json`);
  return state;
}

const STORED_UNLINKED_NAMES = [
  ['a stored session keeps the comment after the first title', 'stored_unlinked_comment_after_first_title'],
  ['a stored session keeps the comment after editing the last title', 'stored_unlinked_comment_after_last_title_edit'],
  ['a stored session keeps the comment after editing the middle title', 'stored_unlinked_comment_after_middle_title_edit'],
] as const;

const MAIN_STATE_CASES = WORD_TOC_SHAPES.flatMap(([name, build]) => {
  const shape = name.split(' ')[0]!;
  const toc = build();
  return WORD_TOC_EDITS.slice(0, shape === 'W4' ? 3 : 2).map(([label, plan], index) => {
    const editName = label.split(' ')[0]!;
    const [title, caseName] = (shape === 'W4' ? STORED_UNLINKED_NAMES[index] : undefined) ?? [
      `${shape} opens a session produced by main's code (${editName})`, `mainstate ${shape} ${editName}`,
    ];
    return [title, caseName, shape, editName, toc, plan(toc)!] as const;
  });
});

test.each(MAIN_STATE_CASES)('%s', async (_title, caseName, shape, name, toc, edit) => {
  const state = await mainState(shape);
  const source = new Uint8Array(Buffer.from(state.docx, 'base64'));
  const update = new Uint8Array(Buffer.from(state.update, 'base64'));
  const expected = editedWordTocXml(toc, edit, await readDocument(source));
  const session = await createYrsSession();
  try {
    session.openDocx(source, false);
    session.loadState(update);
    const applied = applyWordTocEdit(session, edit);
    const { bytes } = await saveYrsDocx(session, { updateModifiedDate: false });
    assertSpanningSave(await readDocument(bytes), expected, caseName);
    expect(applied).toBe(true);
    if (shape === 'W4' && name === 'E3') {
      await expectMatchingUnlinkedProjection();
    }
  } finally {
    session.destroy();
  }
}, 30_000);

async function expectMatchingUnlinkedProjection(): Promise<void> {
  const toc = wordToc({ code: 'TOC', unlinked: true });
  const source = await packageBytes(wordTocXml(toc));
  const native = await createYrsSession();
  try {
    const projected = await createYrsSession();
    try {
      native.openDocx(source, true);
      projected.openDocx(source, false);
      documentToYrs(projected, await parseDocx(source, { preloadFonts: false }));
      const firstTitle = (session: YrsSession) => session.storySegments('body').flatMap(segment =>
        segment.kind === 'text' && segment.attributes.fieldResult ? [segment.text] : []).join('');
      expect(firstTitle(projected)).toBe(firstTitle(native));
      expect(firstTitle(native)).toContain('ENTRYONE');
    } finally {
      projected.destroy();
    }
  } finally {
    native.destroy();
  }
}

test('deleting the last TOC entry closes the field after the last surviving entry', async () => {
  const end = fieldChar('end');
  const second = entry(2, ENTRY_TWO, PAGE_MARKERS[1]);
  const third = entry(3, 'ENTRYTHREE', PAGE_MARKERS[2]);
  const source = body('TA');
  const expected = source.replace(paragraph(third + end), '').replace(second, second + end);
  const xml = await editAndSave(source, session => {
    const { paraId } = session.paragraphs('body')[2]!;
    const result = session.applyEdits({
      expectVersion: session.version(),
      steps: [{ op: 'deleteParagraphs', story: 'body', firstParaId: paraId, lastParaId: paraId }],
    });
    expect(result.ok).toBe(true);
  });
  expect(inspect(xml)).toEqual(inspect(expected));
}, 30_000);

test('merging an end-only paragraph carries the TOC end into the preceding entry', async () => {
  const end = fieldChar('end');
  const third = entry(3, 'ENTRYTHREE', PAGE_MARKERS[2]);
  const source = body('TB');
  const expected = source.replace(paragraph(end), '').replace(third, third + end);
  const xml = await editAndSave(source, session => {
    session.mergeParagraphs('body', session.paragraphs('body')[2]!.paraId);
  });
  expect(inspect(xml)).toEqual(inspect(expected));
}, 30_000);

for (const wrapper of ['ins', 'del']) {
  test(`preserves the boundary and metadata of a marker-only ${wrapper} after an edit`, async () => {
    const boundary = `<w:${wrapper} w:id="17" w:author="Reviewer" w:date="2000-01-01T00:00:00Z">${fieldChar('end')}</w:${wrapper}>`;
    const source = body('S1').replace(fieldChar('end'), boundary);
    for (const fullWrite of [false, true]) {
      const xml = await editAndSave(source, session => insertInText(session, 'SPANEND', 1, ' INSIDE '), fullWrite);
      expect(inspect(xml)).toEqual(inspect(source.replace('SPANEND', 'S INSIDE PANEND')));
      const revision = xml.match(new RegExp(`<w:${wrapper}\\b([^>]*)>([\\s\\S]*?)</w:${wrapper}>`));
      expect(revision?.[1]).toContain('w:id="17"');
      expect(revision?.[1]).toContain('w:author="Reviewer"');
      expect(revision?.[1]).toContain('w:date="2000-01-01T00:00:00Z"');
      expect(revision?.[2]).toContain('w:fldCharType="end"');
    }
  }, 30_000);
}

test('edits on both sides of a spanning end keep the end between the edited texts', async () => {
  const source = body('S1');
  const xml = await editAndSave(source, session => {
    insertInText(session, 'SPANEND', 1, ' INSIDE ');
    insertInText(session, 'AFTERSPAN', 1, ' OUTSIDE ');
  }, true);
  expect(inspect(xml)).toEqual(inspect(source.replace('SPANEND', 'S INSIDE PANEND').replace('AFTERSPAN', 'A OUTSIDE FTERSPAN')));
}, 30_000);

test('consecutive spans keep intervening text outside both fields', async () => {
  const source = [
    fieldStart('QUOTE FIRST') + textRun('FIRST'),
    textRun('FIRSTEND') + fieldChar('end') + textRun('BETWEEN') + fieldStart('QUOTE SECOND') + textRun('SECOND'),
    textRun('SECONDEND') + fieldChar('end') + textRun('OUTSIDE'),
  ].map(paragraph).join('');
  for (const fullWrite of [false, true]) {
    const xml = await editAndSave(source, session => insertInText(session, 'BETWEEN', 1, ' EDIT '), fullWrite);
    expect(inspect(xml)).toEqual(inspect(source.replace('BETWEEN', 'B EDIT ETWEEN')));
  }
}, 30_000);

test('editing both sides of a TOC entry keeps the PAGEREF result after the label', async () => {
  const source = body('TA');
  for (const fullWrite of [false, true]) {
    for (const [label, page] of [['ENTRYONE', 'PAGEONE'], ['ENTRYTWO', 'PAGETWO']] as const) {
      const xml = await editAndSave(source, session => {
        insertInText(session, label, 1, ' LONGER LABEL ');
        insertInText(session, page, 1, ' NEW PAGE ');
      }, fullWrite);
      const expected = source.replace(label, label[0] + ' LONGER LABEL ' + label.slice(1))
        .replace(page, page[0] + ' NEW PAGE ' + page.slice(1));
      expect(inspect(xml)).toEqual(inspect(expected));
    }
  }
}, 30_000);

test('editing numbered first and last TOC entries keeps their PAGEREF instructions', async () => {
  const entries = [['ENTRYONE', 'PAGEONE'], ['ENTRYTHREE', 'PAGETHREE']] as const;
  const tab = '<w:r><w:tab/></w:r>';
  const source = body('TA')
    .replace(textRun('ENTRYONE'), textRun('1') + tab + textRun('ENTRYONE'))
    .replace(textRun('ENTRYTHREE'), textRun('3') + tab + textRun('ENTRYTHREE'));
  for (const fullWrite of [false, true]) {
    for (const [label, page] of entries) {
      const xml = await editAndSave(source, session => {
        insertInText(session, label, 1, ' LONGER TITLE ');
        insertInText(session, page, 1, ' NEW PAGE ');
      }, fullWrite);
      const expected = source.replace(label, label[0] + ' LONGER TITLE ' + label.slice(1))
        .replace(page, page[0] + ' NEW PAGE ' + page.slice(1));
      expect(inspect(xml)).toEqual(inspect(expected));
      const link = [...xml.matchAll(/<w:hyperlink\b[^>]*>([\s\S]*?)<\/w:hyperlink>/g)]
        .find(([, content]) => content?.includes('LONGER TITLE'))?.[1] ?? '';
      expect(link.match(/<w:tab\b/g) ?? []).toHaveLength(2);
      expect(link.lastIndexOf('<w:tab')).toBeLessThan(link.indexOf('w:fldCharType="begin"'));
    }
  }
}, 30_000);

test('a spanning QUOTE keeps a complete nested field in its edited opening result', async () => {
  const source = body('TA').replace('TOC \\o "1-3" \\h \\z \\u', 'QUOTE');
  for (const fullWrite of [false, true]) {
    const xml = await editAndSave(source, session => insertInText(session, 'ENTRYONE', 1, ' EDIT '), fullWrite);
    expect(inspect(xml)).toEqual(inspect(source.replace('ENTRYONE', 'E EDIT NTRYONE')));
  }
}, 30_000);

test('a PAGEREF separator can share the source run with its page text', async () => {
  const source = body('TA').replace(
    fieldChar('separate') + textRun(PAGE_MARKERS[0]),
    '<w:r><w:fldChar w:fldCharType="separate"/><w:t>PAGEONE</w:t></w:r>',
  );
  for (const fullWrite of [false, true]) {
    const xml = await editAndSave(source, session => {
      insertInText(session, 'ENTRYONE', 1, ' EDIT ');
      insertInText(session, 'PAGEONE', 1, ' EDIT ');
    }, fullWrite);
    const expected = source.replace('ENTRYONE', 'E EDIT NTRYONE').replace('PAGEONE', 'P EDIT AGEONE');
    expect(inspect(xml)).toEqual(inspect(expected));
  }
}, 30_000);

test('nested spanning opens keep the legacy full write', async () => {
  const end = fieldChar('end');
  const outer = fieldStart('QUOTE OUTER') + textRun('OUTER');
  const inner = fieldStart('QUOTE INNER') + textRun('INNER');
  const source = [outer + inner, textRun('INNEREND') + end, textRun('OUTEREND') + end + textRun('AFTER')]
    .map(paragraph).join('');
  const expected = [outer + end, textRun('INNEREND'), textRun('OUTEREND') + textRun('AFTER edited')]
    .map(paragraph).join('');
  const xml = await editAndSave(source, session => insertInText(session, 'AFTER', 5, ' edited'), true);
  expect(inspect(xml)).toEqual(inspect(expected));
}, 30_000);

test('cross-container closes keep the legacy opening paragraph', async () => {
  const end = fieldChar('end');
  const outer = fieldStart('QUOTE OUTER') + textRun('OUTER');
  const nested = fieldStart('QUOTE LOCAL') + textRun('LOCAL') + end;
  const cell = `<w:tbl><w:tr><w:tc>${paragraph(end + textRun('CELL'))}</w:tc></w:tr></w:tbl>`;
  const box = paragraph(`<w:r><w:pict><v:shape><v:textbox><w:txbxContent>${
    paragraph(end + textRun('BOX'))
  }</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`);
  const control = blockSdt(paragraph(end + textRun('CONTROL')));
  for (const [container, text] of [[cell, 'CELL'], [box, 'BOX'], [control, 'CONTROL']] as const) {
    const source = paragraph(outer + nested) + container + paragraph(textRun('AFTER'));
    const xml = await editAndSave(source, session => insertInText(session, 'AFTER', 5, ' edited'), true);
    expect(inspect(xml.slice(0, xml.indexOf('</w:p>') + '</w:p>'.length))).toEqual(inspect(paragraph(outer + end)));
    expect(xml).toContain(`>${text}<`);
    expect(xml).toContain('>AFTER edited<');
  }
}, 30_000);

test('a shared end and begin run keeps the legacy opening paragraph', async () => {
  const first = fieldStart('QUOTE FIRST') + textRun('FIRST');
  const shared = '<w:r><w:fldChar w:fldCharType="end"/><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText>QUOTE SECOND</w:instrText></w:r>' + fieldChar('separate') + textRun('SECOND');
  const source = [first, shared, textRun('LAST') + fieldChar('end')].map(paragraph).join('');
  const xml = await editAndSave(source, session => insertInText(session, 'LAST', 4, ' edited'), true);
  const opening = xml.slice(0, xml.indexOf('</w:p>') + '</w:p>'.length);
  expect(inspect(opening)).toEqual(inspect(paragraph(first + fieldChar('end'))));
  expect(xml).toContain('>SECOND<');
  expect(xml).toContain('>LAST edited<');
}, 30_000);

test('both seeders keep the projection id after a same-paragraph raw end', async () => {
  const end = fieldChar('end');
  const raw = `<w:hyperlink w:anchor="top">${fieldChar('begin')}</w:hyperlink>${end}`;
  const next = fieldStart('QUOTE NEXT') + entry(1, 'ENTRYONE', PAGE_MARKERS[0]) + end;
  const source = await packageBytes(paragraph(raw + next) + paragraph(textRun('AFTER')));
  const parsed = await parseDocx(source, { preloadFonts: false });
  for (const seeder of ['native', 'projected'] as const) {
    const session = await createYrsSession();
    try {
      if (seeder === 'native') session.seedFromDocx(source);
      else documentToYrs(session, parsed);
      const nextField = session.storySegments('body').find(segment =>
        segment.kind === 'embed' && segment.embedKind === 'field' && segment.payload.instruction === 'QUOTE NEXT'
      );
      expect(nextField?.kind === 'embed' && nextField.payload.resultProjection).toMatchObject({ id: 2 });
      expect(session.resolveRawMarks()).toEqual([]);
    } finally {
      session.destroy();
    }
  }
}, 30_000);

test('finding s1 keeps a tracked insertion in the opening unlinked title', async () => {
  const source = wordTocXml(wordToc({ code: 'TOC', unlinked: true }));
  const author: YrsAuthor = { name: 'Reviewer', date: '2000-01-01T00:00:00Z' };
  let revisionId: string | null = null;
  const xml = await editAndSave(source, session => {
    const match = session.searchText('ENTRYONE', { caseSensitive: true })[0];
    if (match) {
      const at = { story: match.story, paraId: match.paraId, offset: match.start + 1 };
      revisionId = session.insertText(at, 'ADDED', author).revisionId;
    }
  }, true);
  const insertion = xml.match(/<w:ins\b([^>]*)>[\s\S]*?<\/w:ins>/)?.[0];
  const id = insertion && attribute(insertion, 'w:id');
  const expected = source.replace(textRun('ENTRYONE'),
    textRun('E') + `<w:ins w:id="${id ?? 'missing'}">${textRun('ADDED')}</w:ins>` + textRun('NTRYONE'));
  assertSpanningSave(xml, expected, 'finding_s1_tracked_unlinked_title');
  expect(revisionId).not.toBeNull();
  expect(id).toMatch(/^\d+$/);
  expect(insertion).toContain('w:author="Reviewer"');
  expect(insertion).toContain('w:date="2000-01-01T00:00:00Z"');
}, 30_000);

const LINKED_TITLE_CASES = ['text separator', 'added tab'] as const;

test.each(LINKED_TITLE_CASES.map(kind => [kind] as const))(
  'finding s4 keeps the linked page reference with a %s',
  async kind => {
    const tab = '<w:r><w:tab/></w:r>';
    const source = kind === 'text separator' ? body('TA').replace(tab, textRun('...')) : body('TA');
    const added = kind === 'added tab' ? '\t' : 'EDIT';
    const xml = await editAndSave(source, session => insertInText(session, 'ENTRYONE', 1, added), true);
    const expected = source.replace(textRun('ENTRYONE'), kind === 'added tab'
      ? textRun('E') + tab + textRun('NTRYONE')
      : textRun('EEDITNTRYONE'));
    assertSpanningSave(xml, expected, `finding_s4_${kind.replace(' ', '_')}`);
  },
  30_000,
);

const UNLINKED_NO_PAGE_CASES = (['native', 'projected'] as const).flatMap(seeder =>
  (['first title', 'delete continuation'] as const).map(edit => [seeder, edit] as const));

test.each(UNLINKED_NO_PAGE_CASES)('finding s5 %s loader supports %s without level-one pages', async (seeder, edit) => {
  const toc = wordToc({ code: 'TOC \\n "1-1"', unlinked: true, omitLevelOnePages: true });
  const plan: WordTocEdit = edit === 'first title'
    ? { kind: 'text', story: toc.story, marker: toc.titles[0]! }
    : { kind: 'delete', story: toc.story, index: 1 };
  await assertWordTocSave(toc, plan, `finding_s5_${seeder}_${edit.replace(' ', '_')}`, seeder);
}, 30_000);

test('deleting the last entry and following paragraph closes the TOC in the last surviving entry', async () => {
  const toc = wordToc({ code: 'TOC', inlineEnd: true, unlinked: true });
  const source = wordTocXml(toc);
  const expected = paragraph(toc.paragraphs[0]!) + paragraph(toc.paragraphs[1]! + fieldChar('end'));
  let applied = false;
  const xml = await editAndSave(source, session => {
    const story = 'body';
    const [, , first, last] = session.paragraphs(story);
    applied = session.applyEdits({
      expectVersion: session.version(),
      steps: [{ op: 'deleteParagraphs', story, firstParaId: first!.paraId, lastParaId: last!.paraId }],
    }).ok;
  });
  assertSpanningSave(xml, expected, 'delete_last_entry_and_following_paragraph_carries_end');
  expect(applied).toBe(true);
}, 30_000);

test('finding n4 editing a projected title keeps the simple field and page reference', async () => {
  const simple = `<w:fldSimple w:instr="DATE">${textRun('TODAY')}</w:fldSimple>`;
  const source = body('TA').replace('</w:hyperlink>', '</w:hyperlink>' + simple);
  const xml = await editAndSave(source, session => insertInText(session, 'ENTRYONE', 1, ' EDIT '), true);
  assertSpanningSave(xml, source.replace('ENTRYONE', 'E EDIT NTRYONE'), 'finding_n4_projected_simple_field');
  expect(xml).toMatch(/<w:fldSimple\b[^>]*w:instr="DATE"[^>]*>[\s\S]*?TODAY[\s\S]*?<\/w:fldSimple>/);
  expect(xml).toContain('w:anchor="_Toc1"');
}, 30_000);

test('finding n9 structured export keeps cached title runs in a code-only projection', async () => {
  const toc = wordToc({ code: 'TOC', unlinked: true });
  const codeChild = `<w:hyperlink w:anchor="code">${textRun('CODE')}</w:hyperlink>`;
  toc.paragraphs[0] = toc.paragraphs[0]!.replace(fieldChar('separate'), codeChild + fieldChar('separate'));
  const source = wordTocXml(toc);
  let title: string | undefined;
  const xml = await editAndSave(source, session => {
    const ops: YrsRawOp[] = [];
    let index = 0;
    let owner = false;
    for (const segment of session.storySegments('body')) {
      const len = segment.kind === 'text' ? segment.text.length : 1;
      const mark = segment.attributes.fieldResult as { index: number } | undefined;
      if (mark && mark.index >= 0) ops.unshift({ op: 'delete', index, len });
      if (segment.kind === 'embed' && segment.embedKind === 'field' && segment.payload.fieldType === 'TOC') {
        const projection = segment.payload.resultProjection as { id: number; children: { index: number }[] } | undefined;
        if (projection?.children.some(child => child.index < 0)) {
          owner = true;
          const children = projection.children.filter(child => child.index < 0);
          ops.unshift({ op: 'setEmbedAttr', index, key: 'resultProjection', value: { ...projection, children } });
        }
        break;
      }
      index += len;
    }
    if (owner) {
      session.applyRawOps('body', ops);
      const exported = session.exportStructured({ revisionView: 'accepted' });
      const first = exported.ok ? exported.content.stories.find(story => story.story === 'body')?.blocks[0] : undefined;
      const field = first && 'paragraph' in first
        ? first.paragraph.inlines.find(inline => inline.kind === 'field' && inline.fieldType === 'TOC')
        : undefined;
      if (field?.kind === 'field' && field.cachedResult.kind === 'inline') {
        title = field.cachedResult.inlines.flatMap(inline => inline.kind === 'text' ? [inline.text] : []).join('');
      }
    }
    insertInText(session, ENTRY_TWO, 1, ' EDIT ');
  }, true);
  assertSpanningSave(xml, source.replace(ENTRY_TWO, 'E EDIT NTRYTWO'), 'finding_n9_code_only_projection');
  expect(title).toContain('ENTRYONE');
}, 30_000);
