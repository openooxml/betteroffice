import { beforeAll, describe, expect, it } from 'bun:test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, rezipContainer, unzipContainer } from '../wasm/opc';
import { preloadParseWasm } from '../wasm/parse';
import { residentWorkerFactory, type InProcessResidentWorker } from './__fixtures__/residentWorker';
import { hostSaveMetadata } from './editorSave';
import { createYrsSession, decodeDocxHostJson, saveYrsDocx, type YrsSession, type YrsStorySegment } from './index';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';
import type { DocxCommandEnvelope, DocxCommandInitRequest } from './residentCommandProtocol';

const GEN = (dir: string) => new Uint8Array(readFileSync(resolve(import.meta.dir, `../wasm/generated/${dir}`)));
const FIXTURE_PATH = resolve(import.meta.dir, '../../../../crates/docx-edit/tests/fixtures/structured-export/principal.docx');
const FONT_PATH = resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const BODY_PART = 'word/document.xml';
const BREAK_ID = '00000019';
const NEXT_ID = '00000016';
const OTHER_ID = '00000017';
const TYPED = 'X';
const PAGE_BREAK = /<w:br\b[^>]*w:type="page"[^>]*\/>/g;
const COMMAND_SESSION = 'page-break-beside-inline-content';
const TIMEOUT = 60_000;
const enc = new TextEncoder();
const dec = new TextDecoder();

type Parts = Record<string, Uint8Array>;
type Route = 'worker' | 'loader';
type Kind = { name: string; xml: string; marker: string; result?: string };
type Position = {
  name: string;
  content: (child: string) => string;
  breaks: number;
  order: 'break-first' | 'child-first';
  text?: (word: (needle: string) => string, kind: Kind) => string;
};
type Edit = {
  name: string;
  forced: boolean;
  target?: { id: string; needle: string; at: number };
  text?: string;
  remove?: number;
  resolve?: { accept: boolean; kind: 'insertion' | 'deletion' };
};
type BreakType = 'page' | 'column';
type Place = 'right' | 'mains';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const BREAK_RUN = '<w:r><w:br w:type="page"/></w:r>';

const KINDS: Kind[] = [
  {
    name: 'a complex field',
    xml: '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref1 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Res</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
    marker: 'w:fldCharType="begin"',
    result: 'Res',
  },
  {
    name: 'a simple field',
    xml: '<w:fldSimple w:instr=" REF _Ref1 \\h "><w:r><w:t>Simple</w:t></w:r></w:fldSimple>',
    marker: '<w:fldSimple',
  },
  {
    name: 'a hyperlink',
    xml: '<w:hyperlink w:anchor="_Ref1"><w:r><w:t>Link</w:t></w:r></w:hyperlink>',
    marker: '<w:hyperlink',
    result: 'Link',
  },
  {
    name: 'an inline content control',
    xml: '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:r><w:t>Ctrl</w:t></w:r></w:sdtContent></w:sdt>',
    marker: '<w:sdt>',
  },
  {
    name: 'a tracked insertion',
    xml: '<w:ins w:id="11" w:author="Writer" w:date="2026-10-01T00:00:00Z"><w:r><w:t>Ins</w:t></w:r></w:ins>',
    marker: '<w:ins ',
  },
  {
    name: 'a tracked deletion',
    xml: '<w:del w:id="12" w:author="Writer" w:date="2026-10-01T00:00:00Z"><w:r><w:delText>Del</w:delText></w:r></w:del>',
    marker: '<w:del ',
  },
  {
    name: 'an equation',
    xml: '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>',
    marker: '<m:oMath',
  },
];

const BEFORE: Position = {
  name: 'in its own run before',
  content: (child) => `${run('Head')}${BREAK_RUN}${run('Tail')}${child}`,
  breaks: 1,
  order: 'break-first',
};
const AFTER: Position = {
  name: 'in its own run after',
  content: (child) => `${child}${run('Head')}${BREAK_RUN}${run('Tail')}`,
  breaks: 1,
  order: 'child-first',
};
const AT_START: Position = {
  name: 'opening the paragraph before',
  content: (child) => `${BREAK_RUN}${run('Head')}${child}`,
  breaks: 1,
  order: 'break-first',
};
const TWO_BREAKS: Position = {
  name: 'twice, each in its own run, before',
  content: (child) => `${run('Head')}${BREAK_RUN}${run('Mid')}${BREAK_RUN}${run('Tail')}${child}`,
  breaks: 2,
  order: 'break-first',
};
const IN_RUN: Position = {
  name: 'inside a run between text, before',
  content: (child) => `<w:r><w:t xml:space="preserve">Head</w:t><w:br w:type="page"/><w:t xml:space="preserve">Tail</w:t></w:r>${child}`,
  breaks: 1,
  order: 'break-first',
};
const RIGHT_BEFORE: Position = {
  name: 'in its own run directly before',
  content: (child) => `${run('Head')}${BREAK_RUN}${child}${run('Tail')}`,
  breaks: 1,
  order: 'break-first',
  text: (word, kind) => `${word('Head')}[PB]${kind.result ?? ''}${word('Tail')}`,
};
const KB3: Position = {
  name: 'in its own run between A and B before',
  content: () => {
    const source = fixtureDoc.match(new RegExp(`<w:p w14:paraId="${BREAK_ID}">[\\s\\S]*?<\\/w:p>`))?.[0];
    if (!source) throw new Error(`the fixture has no paragraph ${BREAK_ID}`);
    const field = source.slice(source.indexOf('<w:r><w:fldChar w:fldCharType="begin"/>'), source.lastIndexOf('</w:p>'));
    return `<w:r><w:t>A</w:t></w:r>${BREAK_RUN}<w:r><w:t>B</w:t></w:r>${field}`;
  },
  breaks: 1,
  order: 'break-first',
  text: (word) => `${word('A')}[PB]${word('B')}`,
};

const FORMATTED_RESULT = '<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:tab/><w:t>result</w:t></w:r>';
const FORMATTED_FIELD: Kind = {
  name: 'a complex field whose result is a bold run and a tabbed run',
  xml: `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref1 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${FORMATTED_RESULT}<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
  marker: 'w:fldCharType="begin"',
  result: 'Boldresult',
};
const FORMATTED_LINK: Kind = {
  name: 'a hyperlink holding a bold run and a tabbed run',
  xml: `<w:hyperlink w:anchor="_Ref1">${FORMATTED_RESULT}</w:hyperlink>`,
  marker: '<w:hyperlink',
  result: 'Boldresult',
};
const PRINCIPAL_FIELD: Kind = {
  name: "principal.docx's REF field",
  xml: '',
  marker: 'w:fldCharType="begin"',
  result: 'Boldresult',
};

const NO_EDIT: Edit = { name: 'no edit', forced: false };
const EDIT_BEFORE: Edit = { name: 'the paragraph edited before the break', forced: false, target: { id: BREAK_ID, needle: 'Head', at: 2 } };
const EDIT_AFTER: Edit = { name: 'the paragraph edited after the break', forced: false, target: { id: BREAK_ID, needle: 'Tail', at: 2 } };
const EDIT_OTHER: Edit = { name: 'another paragraph edited on a forced-full document', forced: true, target: { id: OTHER_ID, needle: 'Section two', at: 11 } };
const EDIT_AFTER_LEADING: Edit = { name: 'the paragraph edited after the break', forced: false, target: { id: BREAK_ID, needle: 'Head', at: 2 } };
const EDIT_AT_BREAK: Edit = { name: 'the paragraph edited right at the break', forced: false, target: { id: BREAK_ID, needle: 'Head', at: 4 } };
const EDIT_END_OF_A: Edit = { name: "the paragraph edited at the end of 'A', right at the break", forced: false, target: { id: BREAK_ID, needle: 'A', at: 1 } };

const ROUTES: Route[] = ['worker', 'loader'];

let fixtureParts: Parts;
let fixtureDoc: string;
let font: Uint8Array;
let opaque = '';
let startWorker: (clientId?: number) => InProcessResidentWorker;
let clientId = 84100;

function clone(parts: Parts): Parts {
  return Object.fromEntries(Object.entries(parts).map(([key, value]) => [key, Uint8Array.from(value)]));
}

function editText(parts: Parts, path: string, map: (xml: string) => string): void {
  const bytes = parts[path];
  if (bytes) parts[path] = enc.encode(map(dec.decode(bytes)));
}

function dropParts(parts: Parts, paths: string[]): void {
  for (const path of paths) delete parts[path];
  editText(parts, 'word/_rels/document.xml.rels', (xml) =>
    xml.replace(/<Relationship\b[^>]*\/>/g, (entry) => {
      if (entry.includes('TargetMode="External"')) return entry;
      const target = entry.match(/Target="([^"]+)"/)?.[1];
      return target && paths.includes(`word/${target}`) ? '' : entry;
    }));
  editText(parts, '[Content_Types].xml', (xml) =>
    xml.replace(/<Override\b[^>]*\/>/g, (entry) => (paths.some((path) => entry.includes(`PartName="/${path}"`)) ? '' : entry)));
}

function built(paragraph: string, forced: boolean): Uint8Array {
  const parts = clone(fixtureParts);
  dropParts(parts, ['word/numbering.xml', 'word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml', 'word/media/logo.png']);
  editText(parts, 'word/_rels/document.xml.rels', (xml) =>
    xml.replace(/<Relationship\b[^>]*\/>/g, (entry) => (['rIdImage', 'rIdChart', 'rIdLink'].some((id) => entry.includes(`Id="${id}"`)) ? '' : entry)));
  const root = fixtureDoc.slice(0, fixtureDoc.indexOf('<w:body>'));
  const finalSect = fixtureDoc.slice(fixtureDoc.lastIndexOf('<w:sectPr>'), fixtureDoc.lastIndexOf('</w:body>'));
  const body =
    `<w:p w14:paraId="00000001">${run('Intro')}</w:p>` +
    (forced ? opaque : '') +
    `<w:p w14:paraId="${BREAK_ID}">${paragraph}</w:p>` +
    `<w:p w14:paraId="${NEXT_ID}">${run('End of section one')}</w:p>` +
    `<w:p w14:paraId="${OTHER_ID}">${run('Section two')}</w:p>`;
  parts[BODY_PART] = enc.encode(`${root}<w:body>${body}${finalSect}</w:body></w:document>`);
  return rezipContainer(parts);
}

type SavedParagraph = { xml: string; text: string };

const flowBreak = (type: BreakType) => new RegExp(`<w:br\\b[^>]*w:type="${type}"[^>]*\\/>`, 'g');
const LINE_BREAK = '<w:br(?:\\s+w:type="textWrapping")?\\s*\\/>';

function savedParagraph(xml: string, id: string, type: BreakType = 'page'): SavedParagraph {
  const body = xml.slice(xml.indexOf('<w:body'), xml.lastIndexOf('</w:body>'));
  const found = (body.match(/<w:p(?:\s[^>]*?)?\/>|<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) ?? [])
    .find((p) => (p.match(/^<w:p\b[^>]*?paraId="([0-9A-Fa-f]+)"/)?.[1] ?? '').toUpperCase() === id);
  if (!found) throw new Error(`the saved body has no paragraph ${id}`);
  const text = [...found.matchAll(new RegExp(`<w:t(?:\\s[^>]*)?>([^<]*)<\\/w:t>|<w:br\\b[^>]*w:type="${type}"[^>]*\\/>`, 'g'))]
    .map((m) => (m[1] !== undefined ? m[1] : '[PB]')).join('');
  return { xml: found, text };
}

function decodeXmlText(text: string): string {
  return text
    .replace(/&#x([0-9A-Fa-f]+);/g, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, digits: string) => String.fromCodePoint(Number(digits)))
    .replaceAll('&lt;', '<').replaceAll('&gt;', '>').replaceAll('&quot;', '"').replaceAll('&apos;', "'").replaceAll('&amp;', '&');
}

function flowText(paragraphXml: string, type: BreakType | 'textWrapping'): string {
  const breakPattern = type === 'textWrapping' ? LINE_BREAK : `<w:br\\b[^>]*w:type="${type}"[^>]*\\/>`;
  const token = new RegExp(
    `<w:t(?:\\s[^>]*)?>([^<]*)<\\/w:t>|<w:fldChar\\b[^>]*w:fldCharType="(begin|end)"[^>]*\\/>|(${breakPattern})|` +
    '<w:sym\\b[^>]*\\/>|<w:hyperlink\\b[^>]*>|<\\/w:hyperlink>|<w:fldSimple\\b[^>]*>|<\\/w:fldSimple>', 'g');
  return [...paragraphXml.matchAll(token)].map(([match, text, field, flow]) => {
    if (text !== undefined) return decodeXmlText(text);
    if (field) return field === 'begin' ? '«' : '»';
    if (flow) return type === 'textWrapping' ? '[LB]' : '[PB]';
    if (match.startsWith('<w:sym')) return '[SYM]';
    if (match.startsWith('<w:hyperlink')) return '{';
    if (match === '</w:hyperlink>') return '}';
    return match.startsWith('<w:fldSimple') ? '«' : '»';
  }).join('');
}

function expectRightPlaceOrMains(xml: string, right: RegExp, type: BreakType = 'page'): Place {
  expect(xml.match(flowBreak(type)) ?? []).toHaveLength(1);
  const paragraph = flowText(savedParagraph(xml, BREAK_ID, type).xml, type);
  const next = savedParagraph(xml, NEXT_ID, type).text;
  if (!paragraph.includes('[PB]')) {
    expect(paragraph).toMatch(new RegExp(right.source.replace('\\[PB\\]', ''), right.flags));
    expect(next.startsWith('[PB]')).toBe(true);
    return 'mains';
  }
  expect(next).not.toContain('[PB]');
  expect(paragraph).toMatch(right);
  return 'right';
}

function savedBody(bytes: Uint8Array): string {
  return dec.decode(unzipContainer(bytes)[BODY_PART] ?? new Uint8Array());
}

type Identities = { paragraphIdentities(): { paragraphs: Array<{ ooxmlParaId: string | null; session: { story: string; paraId: string } | null }> } };
type Segmented = { storySegments(story: string): YrsStorySegment[] };

function sessionKey(session: Identities, ooxmlId: string): string {
  const identity = session.paragraphIdentities().paragraphs
    .filter((entry) => entry.session?.story === 'body')
    .find((entry) => entry.ooxmlParaId?.toUpperCase() === ooxmlId);
  if (!identity?.session) throw new Error(`no body paragraph saves with ${ooxmlId}`);
  return identity.session.paraId;
}

function editPoint(session: Identities & Segmented, target: NonNullable<Edit['target']>) {
  const paraId = sessionKey(session, target.id);
  let text = '';
  for (const segment of session.storySegments('body')) {
    if (segment.kind === 'pilcrow') {
      if (segment.paraId === paraId) {
        const at = text.indexOf(target.needle);
        if (at < 0) throw new Error(`no ${target.needle} in ${JSON.stringify(text)}`);
        return { story: 'body', paraId, offset: at + target.at };
      }
      text = '';
    } else if (segment.kind === 'text') {
      text += segment.text;
    } else {
      text += '￼';
    }
  }
  throw new Error(`no paragraph ${target.id} in the body story`);
}

function revisionOf(revisions: readonly { revisionId: string; kind: string }[], kind: 'insertion' | 'deletion'): string {
  const matching = revisions.filter((revision) => revision.kind === kind);
  if (matching.length !== 1) throw new Error(`expected one ${kind}, found ${matching.length}`);
  return matching[0]!.revisionId;
}

async function saveThroughLoader(bytes: Uint8Array, edit: Edit): Promise<Uint8Array> {
  const session: YrsSession = await createYrsSession({ clientId: (clientId += 1) });
  try {
    session.seedFromDocx(Uint8Array.from(bytes), { generation: createHash('sha256').update(bytes).digest('hex').slice(0, 32) });
    if (edit.target) {
      const at = editPoint(session as unknown as Identities & Segmented, edit.target);
      if (edit.resolve) {
        const target = { revisionId: revisionOf(session.listRevisions(), edit.resolve.kind) };
        if (edit.resolve.accept) session.acceptChange(target);
        else session.rejectChange(target);
      } else if (edit.remove) {
        session.deleteRange({ story: 'body', start: { paraId: at.paraId, offset: at.offset }, end: { paraId: at.paraId, offset: at.offset + edit.remove } });
      } else {
        session.replaceRange({ story: 'body', start: { paraId: at.paraId, offset: at.offset }, end: { paraId: at.paraId, offset: at.offset } }, edit.text ?? TYPED);
      }
    }
    return (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
  } finally {
    session.destroy();
  }
}

async function saveThroughWorker(bytes: Uint8Array, edit: Edit): Promise<Uint8Array> {
  const worker = startWorker((clientId += 1));
  const client = new ResidentEngineWorkerClient(worker);
  try {
    const { hostJson } = await client.open(Uint8Array.from(bytes));
    const host = decodeDocxHostJson(hostJson, Uint8Array.from(bytes)).document;
    const engine = worker.sessions[0]!;
    if (edit.target) {
      const at = editPoint(engine as unknown as Identities & Segmented, edit.target);
      const fontId = engine.registerFont(font);
      const measurement = { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' }, fontChains: { 'liberation sans|0|0': [fontId] } };
      const init: DocxCommandInitRequest = {
        session: COMMAND_SESSION, incarnation: 1, epoch: 1, readOnly: false,
        selection: { anchor: at, focus: at },
        fonts: [font],
        layoutInput: JSON.stringify({ bodyStory: 'body', regions: { sections: [{ sectionId: 'main', properties: {} }] }, measurement, renderEnv: {} }),
        layoutWithRegions: true, layoutRevision: 1,
        extras: JSON.stringify({ fontChains: { ...measurement.fontChains, 'calibri|0|0': [fontId] } }),
      };
      await client.commandInit(init);
      const intent = (kind: string, payload: DocxCommandEnvelope['payload']): DocxCommandEnvelope =>
        ({ session: COMMAND_SESSION, incarnation: 1, seq: 1, basis: { epoch: 1, version: 0 }, kind, payload });
      const command = edit.resolve
        ? intent(edit.resolve.accept ? 'acceptRevision' : 'rejectRevision', {
          revisionId: revisionOf(engine.geometryReader.listRevisions(), edit.resolve.kind), expectVersion: engine.proposalEngine.version(),
        })
        : edit.remove
          ? intent('deleteForward', { selection: { anchor: at, focus: { ...at, offset: at.offset + edit.remove } } })
          : intent('insertText', { selection: { anchor: at, focus: at }, text: edit.text ?? TYPED });
      const typed = await client.edit(command);
      expect(typed.outcomes[0]?.status).toBe('applied');
    }
    const saved = await client.save({ comments: [], host: hostSaveMetadata(host) });
    return new Uint8Array(saved.bytes);
  } finally {
    client.destroy();
  }
}

function save(route: Route, bytes: Uint8Array, edit: Edit): Promise<Uint8Array> {
  return route === 'worker' ? saveThroughWorker(bytes, edit) : saveThroughLoader(bytes, edit);
}

function typedInto(edit: Edit, needle: string): string {
  if (edit.target?.id !== BREAK_ID || edit.target.needle !== needle) return needle;
  return `${needle.slice(0, edit.target.at)}${TYPED}${needle.slice(edit.target.at)}`;
}

function expectedText(position: Position, kind: Kind, edit: Edit): string {
  const word = (needle: string) => typedInto(edit, needle);
  if (position.text) return position.text(word, kind);
  if (position === AT_START) return `[PB]${word('Head')}`;
  if (position === TWO_BREAKS) return `${word('Head')}[PB]${word('Mid')}[PB]${word('Tail')}`;
  return `${word('Head')}[PB]${word('Tail')}`;
}

function expectEditLanded(xml: string, edit: Edit): void {
  if (!edit.target) return;
  const { id, needle, at } = edit.target;
  const edited = `${needle.slice(0, at)}${TYPED}${needle.slice(at)}`;
  expect(savedParagraph(xml, id).text.replaceAll('[PB]', '')).toContain(edited);
}

function expectKeptInPlace(xml: string, position: Position, kind: Kind, edit: Edit, type: BreakType = 'page'): void {
  expectEditLanded(xml, edit);
  expect(xml.match(flowBreak(type)) ?? []).toHaveLength(position.breaks);
  const paragraph = savedParagraph(xml, BREAK_ID, type);
  expect(savedParagraph(xml, NEXT_ID, type).text).not.toContain('[PB]');
  const expected = expectedText(position, kind, edit);
  const either = [expected, expected.replace(`${TYPED}[PB]`, `[PB]${TYPED}`), expected.replace(`[PB]${TYPED}`, `${TYPED}[PB]`)];
  expect(paragraph.text).toContain(either.find((text) => paragraph.text.includes(text)) ?? expected);
  const breaks = [...paragraph.xml.matchAll(flowBreak(type))].map((m) => m.index!);
  const child = paragraph.xml.indexOf(kind.marker);
  expect(child).toBeGreaterThan(-1);
  if (position.order === 'break-first') expect(Math.max(...breaks)).toBeLessThan(child);
  else expect(Math.min(...breaks)).toBeGreaterThan(child);
}

function expectMovedToNextParagraph(xml: string): void {
  expect(xml.match(PAGE_BREAK) ?? []).toHaveLength(1);
  expect(savedParagraph(xml, BREAK_ID).text).not.toContain('[PB]');
  expect(savedParagraph(xml, NEXT_ID).text.startsWith('[PB]')).toBe(true);
}

beforeAll(async () => {
  await preloadEditWasm(GEN('edit/docx_edit_bg.wasm'));
  await preloadOpcWasm(GEN('opc/ooxml_opc_bg.wasm'));
  await preloadParseWasm(GEN('parse/docx_parse_bg.wasm'));
  fixtureParts = unzipContainer(new Uint8Array(readFileSync(FIXTURE_PATH)));
  fixtureDoc = dec.decode(fixtureParts[BODY_PART]!);
  opaque = fixtureDoc.match(/<bofx:block\b[\s\S]*?<\/bofx:block>/)?.[0] ?? '';
  if (!opaque) throw new Error('the fixture holds no opaque bofx:block');
  font = new Uint8Array(readFileSync(FONT_PATH));
  startWorker = await residentWorkerFactory();
}, TIMEOUT);

describe('a page break beside inline content keeps its place on save', () => {
  const COMPLEX_FIELD = KINDS[0]!;
  const SIMPLE_FIELD = KINDS[1]!;
  const HYPERLINK = KINDS[2]!;
  const cells: Array<[Kind, Position, Edit]> = [
    ...KINDS.flatMap((kind) => [BEFORE, AFTER].flatMap((position) =>
      [EDIT_BEFORE, EDIT_AFTER, EDIT_OTHER, NO_EDIT].map((edit): [Kind, Position, Edit] => [kind, position, edit]))),
    ...[EDIT_AFTER_LEADING, EDIT_OTHER].map((edit): [Kind, Position, Edit] => [COMPLEX_FIELD, AT_START, edit]),
    ...[EDIT_BEFORE, EDIT_OTHER].map((edit): [Kind, Position, Edit] => [COMPLEX_FIELD, TWO_BREAKS, edit]),
    ...[COMPLEX_FIELD, SIMPLE_FIELD, HYPERLINK].flatMap((kind) =>
      [EDIT_BEFORE, EDIT_AFTER, EDIT_OTHER].map((edit): [Kind, Position, Edit] => [kind, IN_RUN, edit])),
    ...[COMPLEX_FIELD, HYPERLINK].flatMap((kind) =>
      [EDIT_BEFORE, EDIT_AFTER].map((edit): [Kind, Position, Edit] => [kind, RIGHT_BEFORE, edit])),
    ...[FORMATTED_FIELD, FORMATTED_LINK].flatMap((kind) =>
      [EDIT_BEFORE, EDIT_AFTER].map((edit): [Kind, Position, Edit] => [kind, BEFORE, edit])),
    ...[COMPLEX_FIELD, HYPERLINK].map((kind): [Kind, Position, Edit] => [kind, BEFORE, EDIT_AT_BREAK]),
    ...[EDIT_END_OF_A, EDIT_OTHER].map((edit): [Kind, Position, Edit] => [PRINCIPAL_FIELD, KB3, edit]),
  ];

  const typed: Array<[Kind, Position, Edit, BreakType]> = [
    ...cells.map(([kind, position, edit]): [Kind, Position, Edit, BreakType] => [kind, position, edit, 'page']),
    ...KINDS.flatMap((kind) => [BEFORE, AFTER].flatMap((position) =>
      [EDIT_BEFORE, EDIT_AFTER, EDIT_OTHER].map((edit): [Kind, Position, Edit, BreakType] => [kind, position, edit, 'column']))),
  ];

  for (const [kind, position, edit, type] of typed) {
    for (const route of ROUTES) {
      it(`${route} route: a ${type} break ${position.name} ${kind.name}, ${edit.name}, stays between its neighbours`, async () => {
        const bytes = built(position.content(kind.xml).replaceAll('w:type="page"', `w:type="${type}"`), edit.forced);
        const xml = savedBody(await save(route, bytes, edit));
        expectKeptInPlace(xml, position, kind, edit, type);
      }, TIMEOUT);
    }
  }
});

describe('a page break is put back only where its place is certain', () => {
  const COMPLEX_FIELD = KINDS[0]!;
  const BREAK_WITH_Z = '<w:r><w:br w:type="page"/><w:t>Z</w:t></w:r>';
  const SYMBOL_LINK =
    '<w:hyperlink w:anchor="_Ref1"><w:r><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:eastAsia="Wingdings" w:cs="Wingdings"/></w:rPr>' +
    '<w:t>&#xF021;</w:t></w:r></w:hyperlink>';
  const REPEATED_LINK = '<w:hyperlink w:anchor="_Ref1"><w:r><w:t>aaa</w:t></w:r></w:hyperlink>';
  const hyperlinks = (xml: string) => (savedParagraph(xml, BREAK_ID).xml.match(/<w:hyperlink\b/g) ?? []).length;
  const fieldBegins = (xml: string) => (savedParagraph(xml, BREAK_ID).xml.match(/w:fldCharType="begin"/g) ?? []).length;

  const APPEND_TO_Z: Edit = { name: "'X' appended after 'Z' on a forced-full document", forced: true, target: { id: BREAK_ID, needle: 'Z', at: 1 } };
  const typeInLink = (forced: boolean): Edit => ({
    name: `'a' typed between the first two linked letters${forced ? ' on a forced-full document' : ''}`,
    forced, target: { id: BREAK_ID, needle: 'Aaaa', at: 2 }, text: 'a',
  });
  const DELETE_FIRST_FIELD: Edit = { name: 'the first field deleted', forced: false, target: { id: BREAK_ID, needle: '￼', at: 0 }, remove: 1 };
  const resolve = (accept: boolean, kind: 'insertion' | 'deletion'): Edit => ({
    name: `the tracked ${kind} ${accept ? 'accepted' : 'rejected'}`,
    forced: false, target: { id: BREAK_ID, needle: 'A', at: 0 }, resolve: { accept, kind },
  });
  const REVISION_DATE = 'w:author="Writer" w:date="2026-10-01T00:00:00Z"';

  type Cell = { name: string; content: string; edit: Edit; routes: Route[]; check: (xml: string) => void };
  const cells: Cell[] = [
    {
      name: 'a hyperlink holding a symbol-font character',
      content: `<w:r><w:t>A</w:t></w:r>${SYMBOL_LINK}${BREAK_WITH_Z}`,
      edit: APPEND_TO_Z,
      routes: ROUTES,
      check: (xml) => {
        expectRightPlaceOrMains(xml, /^A\{(?:\[SYM\]|)\}\[PB\]ZX$/);
        expect(hyperlinks(xml)).toBe(1);
      },
    },
    ...[false, true].map((forced): Cell => ({
      name: "a hyperlink of repeated letters 'aaa'",
      content: `<w:r><w:t>A</w:t></w:r>${REPEATED_LINK}${BREAK_WITH_Z}`,
      edit: typeInLink(forced),
      routes: ROUTES,
      check: (xml) => {
        expectRightPlaceOrMains(xml, /^A\{aaaa\}\[PB\]Z$/);
        expect(hyperlinks(xml)).toBe(1);
      },
    })),
    {
      name: 'two identical REF fields',
      content: `<w:r><w:t>A</w:t></w:r>${COMPLEX_FIELD.xml}${BREAK_RUN}${COMPLEX_FIELD.xml}<w:r><w:t>Z</w:t></w:r>`,
      edit: DELETE_FIRST_FIELD,
      routes: ROUTES,
      check: (xml) => {
        expectRightPlaceOrMains(xml, /^A\[PB\]«Res»Z$/);
        expect(fieldBegins(xml)).toBe(1);
      },
    },
    ...([[true, 'deletion'], [false, 'insertion']] as const).map(([accept, kind]): Cell => ({
      name: `a tracked ${kind} of 'a' beside a plain 'a'`,
      content: `<w:r><w:t>A</w:t></w:r>` +
        (kind === 'deletion'
          ? `<w:del w:id="12" ${REVISION_DATE}><w:r><w:delText>a</w:delText></w:r></w:del>`
          : `<w:ins w:id="11" ${REVISION_DATE}><w:r><w:t>a</w:t></w:r></w:ins>`) +
        `<w:r><w:t>a</w:t></w:r>${BREAK_WITH_Z}`,
      edit: resolve(accept, kind),
      routes: ROUTES,
      check: (xml) => {
        expectRightPlaceOrMains(xml, /^Aa\[PB\]Z$/);
        expect(savedParagraph(xml, BREAK_ID).xml).not.toMatch(/<w:(?:ins|del)\b/);
      },
    })),
  ];

  for (const { name, content, edit, routes, check } of cells) {
    for (const route of routes) {
      it(`${route} route: a page break beside ${name}, ${edit.name}, is put back between its neighbours or where main puts it`, async () => {
        const bytes = built(content, edit.forced);
        check(savedBody(await save(route, bytes, edit)));
      }, TIMEOUT);
    }
  }
});

describe('a line break beside inline content saves as on main', () => {
  for (const route of ROUTES) {
    it(`${route} route: a line break in its own run before a complex field, the paragraph edited before it, stays inside its run`, async () => {
      const bytes = built(`${run('Head')}<w:r><w:br/></w:r>${run('Tail')}${KINDS[0]!.xml}`, false);
      const xml = savedBody(await save(route, bytes, EDIT_BEFORE));
      const paragraph = savedParagraph(xml, BREAK_ID).xml;
      expect(xml.match(PAGE_BREAK) ?? []).toHaveLength(0);
      expect(xml.match(flowBreak('column')) ?? []).toHaveLength(0);
      expect(paragraph.match(new RegExp(LINE_BREAK, 'g')) ?? []).toHaveLength(1);
      expect(paragraph).toMatch(new RegExp(`<w:r(?:\\s[^>]*)?>(?:(?!<\\/w:r>)[\\s\\S])*?${LINE_BREAK}`));
      expect(flowText(paragraph, 'textWrapping')).toBe('HeXad[LB]Tail«Res»');
      expect(savedParagraph(xml, NEXT_ID).xml).not.toMatch(/<w:br\b/);
    }, TIMEOUT);
  }
});

describe('a page break beside inline content whose width may not match keeps today\'s placement', () => {
  const FIELD_RESULT_LINK =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK \\l "_Ref1" </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Ref1"><w:r><w:t>Link</w:t></w:r></w:hyperlink>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const FIELD_IN_LINK =
    '<w:hyperlink w:anchor="_Ref1"><w:r><w:t>Go</w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:hyperlink>';
  const EDIT_FIELD_RESULT: Edit = { name: 'the field result edited', forced: false, target: { id: BREAK_ID, needle: 'Link', at: 2 } };

  const cells: Array<[string, string, Edit, Route[]]> = [
    ['a field whose result is edited in the session', FIELD_RESULT_LINK, EDIT_FIELD_RESULT, ['loader']],
    ['a field nested in a hyperlink', FIELD_IN_LINK, EDIT_BEFORE, ROUTES],
    ['a field nested in a hyperlink on a forced-full document', FIELD_IN_LINK, EDIT_OTHER, ROUTES],
  ];

  for (const [name, child, edit, routes] of cells) {
    for (const route of routes) {
      it(`${route} route: keeps today's placement when the paragraph holds ${name}, ${edit.name}: the break moves to the next paragraph`, async () => {
        const bytes = built(BEFORE.content(child), edit.forced);
        const xml = savedBody(await save(route, bytes, edit));
        expectMovedToNextParagraph(xml);
      }, TIMEOUT);
    }
  }
});
