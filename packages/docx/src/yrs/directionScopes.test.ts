import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { documentToYrs } from './documentToYrs';
import { saveEditorDocument, type EditorSaveRecord } from './editorSave';
import { createYrsSession, type YrsSession } from './index';
import { savedCommentIds } from './residentCommentProjection';
import { yrsToDocument } from './yrsToDocument';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const SECTION = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

const run = (inner: string) => `<w:r>${inner}</w:r>`;
const t = (value: string) => run(`<w:t xml:space="preserve">${value}</w:t>`);
const dt = (value: string) => run(`<w:delText xml:space="preserve">${value}</w:delText>`);
const dir = (val: string, inner: string) => `<w:dir w:val="${val}">${inner}</w:dir>`;
const bdo = (val: string, inner: string) => `<w:bdo w:val="${val}">${inner}</w:bdo>`;
const INS = (inner: string) => `<w:ins w:id="41" w:author="Grace" w:date="2025-03-04T05:06:07Z">${inner}</w:ins>`;
const DEL = (inner: string) => `<w:del w:id="42" w:author="Grace" w:date="2025-03-04T05:06:07Z">${inner}</w:del>`;
const SDT = (inner: string) => `<w:sdt><w:sdtPr><w:id w:val="1"/><w:tag w:val="cc"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const fldChar = (type: string) => run(`<w:fldChar w:fldCharType="${type}"/>`);

const paragraph = (inner: string) => `<w:p w14:paraId="1A2B3C01" w14:textId="1A2B3C01">${t('Head ')}${inner}${t(' tail')}</w:p>`;
const OTHER = `<w:p w14:paraId="1A2B3C02" w14:textId="1A2B3C02">${t('Other')}</w:p>`;

function docx(inner: string): Uint8Array {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', toBytes(`<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:w14="${W14}"><w:body>${paragraph(inner)}${OTHER}${SECTION}</w:body></w:document>`)],
  ])));
}

const part = (bytes: Uint8Array) => new TextDecoder().decode(unzipContainer(bytes)['word/document.xml']);
const count = (xml: string, pattern: RegExp) => xml.match(pattern)?.length ?? 0;
const body = (xml: string) => xml.slice(xml.indexOf('<w:body>'), xml.indexOf('<w:sectPr'));
const paragraphsOf = (xml: string) => body(xml).match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];

const TOKEN = /<w:(t|delText)\b[^>]*>([^<]*)<\/w:\1>/g;
const render = (xml: string) => [...xml.matchAll(TOKEN)].map(([, , text]) => text).join('');
const deleted = (xml: string) => [...xml.matchAll(/<w:delText\b[^>]*>([^<]*)<\/w:delText>/g)].map(([, text]) => text).join('');

/** Raw content of every `w:<tag>` wrapper whose w:val is `val`, joined in document order. */
function scoped(xml: string, tag: string, val: string): { wrappers: number; raw: string } {
  const matching = [...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*\\bw:val="${val}"[^>]*>([\\s\\S]*?)</w:${tag}>`, 'g'))];
  return { wrappers: matching.length, raw: matching.map(([, inner]) => inner).join('') };
}

const enclosed = (xml: string, outer: string, inner: string) =>
  new RegExp(`<w:${outer}\\b[^>]*>(?:(?!</w:${outer}>)[\\s\\S])*<w:${inner}\\b`).test(xml);

const STRUCTURE = ['hyperlink', 'ins', 'sdt', 'sdtContent', 'fldSimple', 'fldChar', 'instrText'];
const structure = (xml: string) => STRUCTURE.map(tag => [tag, count(xml, new RegExp(`<w:${tag}\\b`, 'g'))]);

const editorRecords = new WeakMap<YrsSession, EditorSaveRecord>();

/** Saves `session` as the editor saves it: projected over `base` with its comments and those made in the session. */
async function savedAsTheEditorSaves(session: YrsSession, base: Document): Promise<Uint8Array> {
  const source = base.package.document.comments ?? [];
  const listed = session.listComments();
  const ids = savedCommentIds(source.map(comment => comment.id), listed);
  const added = listed
    .filter(comment => !source.some(known => String(known.id) === comment.id))
    .map(comment => ({ id: ids.get(comment.id)!, author: comment.author, content: [] }));
  const comments = [...source, ...added];
  const commentIds = new Map([...ids].filter(([id, numeric]) => id !== String(numeric))
    .map(([id, numeric]) => [numeric, id]));
  const record = editorRecords.get(session) ?? { full: false };
  record.commentIds = commentIds;
  editorRecords.set(session, record);
  const projected = yrsToDocument(
    session,
    { ...base, package: { ...base.package, document: { ...base.package.document, comments } } },
    { commentIds },
  );
  return new Uint8Array(await saveEditorDocument(session, projected, comments, record));
}

const sessions: YrsSession[] = [];
let clientId = 78100;

const LOADERS: Array<[string, (bytes: Uint8Array) => Promise<{ session: YrsSession; base: Document }>]> = [
  ['the Rust loader', async bytes => {
    const session = await createYrsSession({ clientId: (clientId += 1) });
    sessions.push(session);
    session.openDocx(bytes, true);
    return { session, base: session.materializeDocx()! };
  }],
  ['the TypeScript loader', async bytes => {
    const session = await createYrsSession({ clientId: (clientId += 1) });
    sessions.push(session);
    const parsed = await parseDocx(bytes.slice().buffer as ArrayBuffer, { preloadFonts: false });
    documentToYrs(session, parsed);
    return { session, base: { ...parsed, originalBuffer: parsed.originalBuffer ?? (bytes.slice().buffer as ArrayBuffer) } };
  }],
];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}, 30_000);
afterEach(() => { sessions.splice(0).forEach(session => session.destroy()); });

type Container = {
  name: string;
  inner: string;
  scopes: Array<[string, string]>;
  text: string;
  embedded?: true;
  refused?: true;
  container?: [string, string];
  holds?: RegExp;
  untouched?: [string, string, string];
};

const CONTAINERS: Container[] = [
  { name: 'a w:dir rtl in a plain paragraph', inner: dir('rtl', t('two')), scopes: [['dir', 'rtl']], text: 'Head two tail' },
  { name: 'a w:bdo ltr in a plain paragraph', inner: bdo('ltr', t('two')), scopes: [['bdo', 'ltr']], text: 'Head two tail' },
  {
    name: 'a w:dir inside a hyperlink',
    inner: `<w:hyperlink w:anchor="target">${dir('rtl', t('two'))}</w:hyperlink>`,
    scopes: [['dir', 'rtl']],
    text: 'Head two tail',
    container: ['hyperlink', 'dir'],
  },
  {
    name: 'a w:dir inside a tracked insertion',
    inner: INS(dir('rtl', t('two'))),
    scopes: [['dir', 'rtl']],
    text: 'Head two tail',
    refused: true,
  },
  {
    name: 'a w:dir inside a tracked deletion',
    inner: DEL(dir('rtl', dt('two'))),
    scopes: [['dir', 'rtl']],
    text: 'Head two tail',
    container: ['del', 'dir'],
  },
  {
    name: 'a w:dir inside a run-level content control',
    inner: SDT(dir('rtl', t('two'))),
    scopes: [['dir', 'rtl']],
    text: 'Head  tail',
    embedded: true,
    container: ['sdtContent', 'dir'],
  },
  {
    name: 'a w:dir inside a simple field',
    inner: `<w:fldSimple w:instr=" PAGE ">${dir('rtl', t('two'))}</w:fldSimple>`,
    scopes: [['dir', 'rtl']],
    text: 'Head  tail',
    embedded: true,
    container: ['fldSimple', 'dir'],
  },
  {
    name: 'a w:dir nested in a w:bdo',
    inner: bdo('ltr', dir('rtl', t('two'))),
    scopes: [['bdo', 'ltr'], ['dir', 'rtl']],
    text: 'Head two tail',
    container: ['bdo', 'dir'],
  },
  {
    name: 'two adjacent w:dir with different values',
    inner: `${dir('rtl', t('two'))}${dir('ltr', t('three'))}`,
    scopes: [['dir', 'rtl']],
    text: 'Head twothree tail',
    untouched: ['dir', 'ltr', 'three'],
  },
  {
    name: 'a w:dir spanning a complex field from separate to end',
    inner: `${fldChar('begin')}${run('<w:instrText xml:space="preserve"> PAGE </w:instrText>')}${dir('rtl', `${fldChar('separate')}${t('two')}${fldChar('end')}`)}`,
    scopes: [['dir', 'rtl']],
    text: 'Head  tail',
    embedded: true,
    holds: /fldCharType="separate"[\s\S]*>two<[\s\S]*fldCharType="end"/,
  },
];

const fieldOrControlText = (session: YrsSession) =>
  session.storySegments('body').filter(segment => segment.kind === 'embed' && JSON.stringify(segment.payload).includes('two'));

for (const container of CONTAINERS) {
  const source = paragraph(container.inner);
  const typedText = container.embedded ? `X${render(source)}` : render(source).replace('two', 'twXo');
  const scopeText = container.embedded ? 'two' : 'twXo';

  describe(container.name, () => {
    for (const [loader, load] of LOADERS) {
      it(`${loader}: the session holds the inner text in order`, async () => {
        const { session } = await load(docx(container.inner));
        expect(session.paragraphs('body').map(entry => entry.text)).toEqual([container.text, 'Other']);
        if (container.embedded) expect(fieldOrControlText(session)).toHaveLength(1);
      });

      if (!container.embedded) {
        it(`${loader}: searching finds the inner text once`, async () => {
          const { session } = await load(docx(container.inner));
          const found = session.searchText('two');
          expect(found).toHaveLength(1);
          expect(found[0]).toMatchObject({ story: 'body', start: 5, end: 8, text: 'two' });
        });
      }

      it(`${loader}: saving without an edit writes the source document part unchanged`, async () => {
        const bytes = docx(container.inner);
        const { session, base } = await load(bytes);
        expect(part(await savedAsTheEditorSaves(session, base))).toBe(part(bytes));
      });

      if (container.refused) {
        it(`${loader}: typing inside the inserted text is refused, as it is without the wrapper`, async () => {
          for (const inner of [container.inner, INS(t('two'))]) {
            const { session } = await load(docx(inner));
            const read = session.readParagraphs({ story: 'body', view: 'accepted' });
            if (!read.ok) throw new Error(read.failure.message);
            const refused = session.applyEdits({
              expectVersion: read.version,
              steps: [{
                op: 'insertText',
                target: { kind: 'search', text: 'tw', within: { kind: 'paragraph', story: 'body', paraId: read.paragraphs[0]!.paraId }, view: 'accepted' },
                at: 'end',
                text: 'X',
              }],
            });
            expect(refused).toMatchObject({ ok: false, failure: { code: 'tracked-revision-conflict' } });
          }
        });
        continue;
      }

      it(`${loader}: typing ${container.embedded ? 'beside' : 'inside'} the scoped text keeps the wrapper and its value`, async () => {
        const bytes = docx(container.inner);
        const { session, base } = await load(bytes);
        if (container.embedded) {
          session.insertText({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 0 }, 'X');
        } else {
          const [match] = session.searchText('two');
          session.insertText({ story: match!.story, paraId: match!.paraId, offset: match!.start + 2 }, 'X');
        }
        const xml = part(await savedAsTheEditorSaves(session, base));
        for (const [tag, val] of container.scopes) {
          const wrapper = scoped(xml, tag, val);
          expect([tag, val, wrapper.wrappers > 0, render(wrapper.raw)]).toEqual([tag, val, true, scopeText]);
        }
        if (container.container) expect(enclosed(xml, ...container.container)).toBe(true);
        if (container.holds) expect(scoped(xml, ...container.scopes[0]!).raw).toMatch(container.holds);
        if (container.untouched) {
          const [tag, val, text] = container.untouched;
          expect(render(scoped(xml, tag, val).raw)).toBe(text);
        }
        expect(render(body(xml))).toBe(`${typedText}Other`);
        expect(deleted(xml).replace('X', '')).toBe(deleted(source));
        expect(structure(xml)).toEqual(structure(source));
        expect(xml).toContain(OTHER);
        if (container.inner.includes('w:anchor="target"')) expect(xml).toContain('w:anchor="target"');
      });
    }
  });
}

describe('editing across the edge of a w:dir', () => {
  const inner = dir('rtl', t('two'));

  const opened = async (load: (bytes: Uint8Array) => Promise<{ session: YrsSession; base: Document }>) => {
    const { session, base } = await load(docx(inner));
    const [match] = session.searchText('two');
    return { session, base, paraId: match!.paraId, start: match!.start };
  };

  for (const [loader, load] of LOADERS) {
    for (const [edge, shift, expected] of [['start', 0, 'Head Xtwo tail'], ['end', 3, 'Head twoX tail']] as const) {
      it(`${loader}: typing at the ${edge} of the scope loses no text and keeps the scope around its text`, async () => {
        const { session, base, paraId, start } = await opened(load);
        session.insertText({ story: 'body', paraId, offset: start + shift }, 'X');
        const xml = part(await savedAsTheEditorSaves(session, base));
        expect(render(body(xml))).toBe(`${expected}Other`);
        const wrapper = scoped(xml, 'dir', 'rtl');
        expect(wrapper.wrappers).toBeGreaterThan(0);
        expect(render(wrapper.raw).replace('X', '')).toBe('two');
        expect(xml).toContain(OTHER);
      });
    }

    for (const [edge, from, to, expected, remaining] of [
      ['start', -2, 1, 'Heawo tail', 'wo'],
      ['end', 2, 5, 'Head twail', 'tw'],
    ] as const) {
      it(`${loader}: deleting a selection across the ${edge} of the scope keeps the scope around the text it still holds`, async () => {
        const { session, base, paraId, start } = await opened(load);
        session.deleteRange({ story: 'body', start: { paraId, offset: start + from }, end: { paraId, offset: start + to } });
        expect(session.paragraphs('body').map(entry => entry.text)).toEqual([expected, 'Other']);
        const xml = part(await savedAsTheEditorSaves(session, base));
        expect(render(body(xml))).toBe(`${expected}Other`);
        const wrapper = scoped(xml, 'dir', 'rtl');
        expect([wrapper.wrappers, render(wrapper.raw)]).toEqual([1, remaining]);
        expect(xml).toContain(OTHER);
      });
    }

    it(`${loader}: splitting the paragraph inside the scope keeps the scope in both halves`, async () => {
      const { session, base, paraId, start } = await opened(load);
      session.splitParagraph({ story: 'body', paraId, offset: start + 2 });
      expect(session.paragraphs('body').map(entry => entry.text)).toEqual(['Head tw', 'o tail', 'Other']);
      const xml = part(await savedAsTheEditorSaves(session, base));
      const [first, second, third] = paragraphsOf(xml);
      expect([render(first!), render(second!), render(third!)]).toEqual(['Head tw', 'o tail', 'Other']);
      expect(render(scoped(first!, 'dir', 'rtl').raw)).toBe('tw');
      expect(render(scoped(second!, 'dir', 'rtl').raw)).toBe('o');
      expect(third).toBe(OTHER);
    });
  }
});
