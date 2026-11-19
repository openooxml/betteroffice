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
type Kind = { name: string; xml: string; marker: string };
type Position = { name: string; content: (child: string) => string; breaks: number; order: 'break-first' | 'child-first' };
type Edit = { name: string; forced: boolean; target?: { id: string; needle: string; at: number } };

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const BREAK_RUN = '<w:r><w:br w:type="page"/></w:r>';

const KINDS: Kind[] = [
  {
    name: 'a complex field',
    xml: '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref1 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Res</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
    marker: 'w:fldCharType="begin"',
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

const NO_EDIT: Edit = { name: 'no edit', forced: false };
const EDIT_BEFORE: Edit = { name: 'the paragraph edited before the break', forced: false, target: { id: BREAK_ID, needle: 'Head', at: 2 } };
const EDIT_AFTER: Edit = { name: 'the paragraph edited after the break', forced: false, target: { id: BREAK_ID, needle: 'Tail', at: 2 } };
const EDIT_OTHER: Edit = { name: 'another paragraph edited on a forced-full document', forced: true, target: { id: OTHER_ID, needle: 'Section two', at: 11 } };
const EDIT_AFTER_LEADING: Edit = { name: 'the paragraph edited after the break', forced: false, target: { id: BREAK_ID, needle: 'Head', at: 2 } };

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

function savedParagraph(xml: string, id: string): SavedParagraph {
  const body = xml.slice(xml.indexOf('<w:body'), xml.lastIndexOf('</w:body>'));
  const found = (body.match(/<w:p(?:\s[^>]*?)?\/>|<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) ?? [])
    .find((p) => (p.match(/^<w:p\b[^>]*?paraId="([0-9A-Fa-f]+)"/)?.[1] ?? '').toUpperCase() === id);
  if (!found) throw new Error(`the saved body has no paragraph ${id}`);
  const text = [...found.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:br\b[^>]*w:type="page"[^>]*\/>/g)]
    .map((m) => (m[1] !== undefined ? m[1] : '[PB]')).join('');
  return { xml: found, text };
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

async function saveThroughLoader(bytes: Uint8Array, edit: Edit): Promise<Uint8Array> {
  const session: YrsSession = await createYrsSession({ clientId: (clientId += 1) });
  try {
    session.seedFromDocx(Uint8Array.from(bytes), { generation: createHash('sha256').update(bytes).digest('hex').slice(0, 32) });
    if (edit.target) {
      const at = editPoint(session as unknown as Identities & Segmented, edit.target);
      session.replaceRange({ story: 'body', start: { paraId: at.paraId, offset: at.offset }, end: { paraId: at.paraId, offset: at.offset } }, TYPED);
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
      const command: DocxCommandEnvelope = {
        session: COMMAND_SESSION, incarnation: 1, seq: 1, basis: { epoch: 1, version: 0 },
        kind: 'insertText', payload: { selection: { anchor: at, focus: at }, text: TYPED },
      };
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

function expectedText(position: Position, edit: Edit): string {
  const word = (needle: string) => typedInto(edit, needle);
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

function expectKeptInPlace(xml: string, position: Position, kind: Kind, edit: Edit): void {
  expectEditLanded(xml, edit);
  expect(xml.match(PAGE_BREAK) ?? []).toHaveLength(position.breaks);
  const paragraph = savedParagraph(xml, BREAK_ID);
  expect(savedParagraph(xml, NEXT_ID).text).not.toContain('[PB]');
  expect(paragraph.text).toContain(expectedText(position, edit));
  const breaks = [...paragraph.xml.matchAll(PAGE_BREAK)].map((m) => m.index!);
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
  const cells: Array<[Kind, Position, Edit]> = [
    ...KINDS.flatMap((kind) => [BEFORE, AFTER].flatMap((position) =>
      [EDIT_BEFORE, EDIT_AFTER, EDIT_OTHER, NO_EDIT].map((edit): [Kind, Position, Edit] => [kind, position, edit]))),
    ...[EDIT_AFTER_LEADING, EDIT_OTHER].map((edit): [Kind, Position, Edit] => [COMPLEX_FIELD, AT_START, edit]),
    ...[EDIT_BEFORE, EDIT_OTHER].map((edit): [Kind, Position, Edit] => [COMPLEX_FIELD, TWO_BREAKS, edit]),
  ];

  for (const [kind, position, edit] of cells) {
    for (const route of ROUTES) {
      it(`${route} route: a page break ${position.name} ${kind.name}, ${edit.name}, stays between its neighbours`, async () => {
        const bytes = built(position.content(kind.xml), edit.forced);
        const xml = savedBody(await save(route, bytes, edit));
        expectKeptInPlace(xml, position, kind, edit);
      }, TIMEOUT);
    }
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
