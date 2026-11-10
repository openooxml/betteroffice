import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeYrs from '@betteroffice/docx-native/yrs';

type Yrs = typeof wasmYrs;
type Session = Awaited<ReturnType<Yrs['createYrsSession']>>;
type Author = { name: string; date: string };
type Edit = (session: Session, author?: Author) => void;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const AUTHOR: Author = { name: 'Ada', date: '2026-01-01T00:00:00Z' };
const SECTION = '<w:sectPr><w:pgSz w:w="7200" w:h="4320"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>';
const text = (value: string) => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
const brk = (kind: string) => `<w:br w:type="${kind}"/>`;
const L = `<w:p w14:paraId="72000000">${text('L')}</w:p>`;
const paragraph = (inner: string) => `<w:p w14:paraId="72000001">${inner}</w:p>`;
const N = `<w:p w14:paraId="72000003">${text('N')}</w:p>`;

const SHAPES: Record<string, (kind: string) => string> = {
  middle: (kind) => L + paragraph(`<w:r><w:t>A</w:t>${brk(kind)}<w:t>C</w:t></w:r>`) + N,
  'middle in three runs': (kind) => L + paragraph(`${text('A')}<w:r>${brk(kind)}</w:r>${text('C')}`) + N,
  end: (kind) => L + paragraph(`<w:r><w:t>A</w:t>${brk(kind)}</w:r>`) + N,
  'own paragraph': (kind) => L + paragraph(`<w:r>${brk(kind)}</w:r>`) + N,
  'before a bare section mark': (kind) => L + paragraph(`<w:r><w:t>A</w:t>${brk(kind)}<w:t xml:space="preserve"> </w:t></w:r>`)
    + `<w:p w14:paraId="72000002"><w:pPr>${SECTION}</w:pPr></w:p><w:p w14:paraId="72000004">${text('B')}</w:p>`,
};

async function docx(body: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}${SECTION}</w:body></w:document>`);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const ids = (session: Session) => session.paragraphs('body').map((entry) => entry.paraId);
const length = (session: Session, index: number) => session.paragraphs('body')[index]!.text.length;
const at = (session: Session, index: number, offset: number) => ({ paraId: ids(session)[index]!, offset });

const EDITS: Record<string, Edit> = {
  'split before the break': (s, a) => { s.splitParagraph({ story: 'body', ...at(s, 1, 1) }, a); },
  'join with the previous paragraph': (s, a) => { s.deleteRange({ story: 'body', start: at(s, 0, 1), end: at(s, 1, 0) }, a); },
  'merge the previous paragraph': (s, a) => { s.mergeParagraphs('body', ids(s)[0]!, a); },
  'paste two paragraphs before the break': (s, a) => {
    s.insertText({ story: 'body', ...at(s, 1, 1) }, 'Q', a);
    const split = s.splitParagraph({ story: 'body', ...at(s, 1, 2) }, a);
    s.insertText({ story: 'body', paraId: split.secondParaId, offset: 0 }, 'R', a);
  },
  'type at the start of the break paragraph': (s, a) => { s.insertText({ story: 'body', ...at(s, 1, 0) }, 'x', a); },
  'type at the end of the break paragraph': (s, a) => { s.insertText({ story: 'body', ...at(s, 1, length(s, 1)) }, 'x', a); },
  'delete from the previous paragraph into the break paragraph': (s, a) => { s.deleteRange({ story: 'body', start: at(s, 0, 0), end: at(s, 1, 0) }, a); },
  'Enter at the start of the break paragraph': (s, a) => { s.splitParagraph({ story: 'body', ...at(s, 1, 0) }, a); },
  'remove the text past the break': (s, a) => { s.deleteRange({ story: 'body', start: at(s, 1, 1), end: at(s, 1, 2) }, a); },
  'delete the break paragraph': (s, a) => { s.deleteRange({ story: 'body', start: at(s, 0, 1), end: at(s, 1, length(s, 1)) }, a); },
  'Delete at the end of the break paragraph': (s, a) => { s.deleteRange({ story: 'body', start: at(s, 1, length(s, 1)), end: at(s, 2, 0) }, a); },
  'split at the end of the break paragraph': (s, a) => { s.splitParagraph({ story: 'body', ...at(s, 1, length(s, 1)) }, a); },
  'merge with the next paragraph': (s, a) => { s.mergeParagraphs('body', ids(s)[1]!, a); },
  'type before the break': (s, a) => { s.insertText({ story: 'body', ...at(s, 1, 1) }, 'x', a); },
  'Enter at the end of the previous paragraph': (s, a) => { s.splitParagraph({ story: 'body', ...at(s, 0, 1) }, a); },
  'join the next paragraph': (s, a) => { s.deleteRange({ story: 'body', start: at(s, 1, length(s, 1)), end: at(s, 2, 0) }, a); },
};

const MODES = ['plain', 'undo and redo', 'tracked', 'remote peer'] as const;
type Mode = typeof MODES[number];

/** [shape, edit, the paragraph `L` stays as its source bytes] */
const ROWS: Array<[string, string, boolean]> = [
  ...['middle', 'middle in three runs', 'before a bare section mark'].flatMap((shape): Array<[string, string, boolean]> => [
    [shape, 'split before the break', true],
    [shape, 'join with the previous paragraph', false],
    [shape, 'merge the previous paragraph', false],
    [shape, 'remove the text past the break', true],
    [shape, 'split at the end of the break paragraph', true],
    [shape, 'delete the break paragraph', false],
    [shape, 'merge with the next paragraph', true],
    [shape, 'type before the break', true],
  ]),
  ['end', 'join with the previous paragraph', false],
  ['end', 'merge the previous paragraph', false],
  ['end', 'paste two paragraphs before the break', true],
  ['end', 'split before the break', true],
  ['end', 'Delete at the end of the break paragraph', true],
  ['end', 'type before the break', true],
  ['end', 'delete the break paragraph', false],
  ['own paragraph', 'join with the previous paragraph', false],
  ['own paragraph', 'merge the previous paragraph', false],
  ['own paragraph', 'delete from the previous paragraph into the break paragraph', false],
  ['own paragraph', 'type at the start of the break paragraph', true],
  ['own paragraph', 'type at the end of the break paragraph', true],
  ['own paragraph', 'Enter at the start of the break paragraph', true],
  ['own paragraph', 'delete the break paragraph', false],
  ['own paragraph', 'Enter at the end of the previous paragraph', false],
  ['own paragraph', 'join the next paragraph', true],
];

const blockBreaks = (session: Session, kind: string) => session.storySegments('body')
  .filter((segment) => segment.kind === 'embed' && segment.embedKind === (kind === 'page' ? 'pageBreak' : 'columnBreak')).length
  + (kind === 'page' ? session.storySegments('body').filter((segment) => segment.kind === 'pilcrow' && segment.properties.pageBreakBeforeRun === true).length : 0);

/** The text that opens each page, read from the story: a page break ends one page and starts the next. */
function pageStarts(session: Session): string[] {
  const starts: string[] = [''];
  let pending = false;
  for (const segment of session.storySegments('body')) {
    if (segment.kind === 'embed' && segment.embedKind === 'pageBreak') { pending = true; continue; }
    if (segment.kind === 'pilcrow') {
      if (segment.properties.pageBreakBeforeRun === true) pending = true;
      continue;
    }
    if (segment.kind !== 'text' || !segment.text.trim()) continue;
    if (pending) { starts.push(''); pending = false; }
    if (!starts[starts.length - 1]) starts[starts.length - 1] = segment.text.trim().slice(0, 3);
  }
  return starts;
}

async function savedXml(yrs: Yrs, session: Session) {
  const saved = await yrs.saveYrsDocx(session, { updateModifiedDate: false });
  return { bytes: saved.bytes, xml: await (await JSZip.loadAsync(saved.bytes)).file('word/document.xml')!.async('string') };
}

async function row(yrs: Yrs, kind: string, shape: string, edit: string, mode: Mode, keepsL: boolean) {
  const bytes = await docx(SHAPES[shape]!(kind));
  const sessions: Session[] = [];
  const open = async (source: Uint8Array, seed: boolean) => {
    const session = await yrs.createYrsSession();
    sessions.push(session);
    session.openDocx(source, seed);
    return session;
  };
  try {
    const host = await open(bytes, true);
    let editor = host;
    if (mode === 'remote peer') {
      editor = await open(bytes, false);
      editor.applyUpdate(host.encodeState());
    }
    EDITS[edit]!(editor, mode === 'tracked' ? AUTHOR : undefined);
    if (mode === 'undo and redo') { editor.undo(); editor.redo(); }
    if (mode === 'remote peer') host.applyUpdate(editor.encodeState());
    const live = blockBreaks(host, kind);
    const saves = [await savedXml(yrs, host), ...(mode === 'remote peer' ? [await savedXml(yrs, editor)] : [])];
    const problems: string[] = [];
    for (const [index, save] of saves.entries()) {
      const who = index === 0 ? 'host' : 'peer';
      const written = (save.xml.match(new RegExp(`w:type="${kind}"`, 'g')) ?? []).length;
      if (written !== live) problems.push(`${who} saved ${written} ${kind} breaks, live ${live}`);
      const reopened = await open(save.bytes, true);
      if (blockBreaks(reopened, kind) !== live) problems.push(`${who} reopened ${blockBreaks(reopened, kind)} ${kind} breaks, live ${live}`);
      if (kind === 'page' && JSON.stringify(pageStarts(reopened)) !== JSON.stringify(pageStarts(host))) {
        problems.push(`${who} reopened pages start ${JSON.stringify(pageStarts(reopened))}, live ${JSON.stringify(pageStarts(host))}`);
      }
      if (keepsL && mode !== 'tracked' && !save.xml.includes(L)) problems.push(`${who} rewrote the untouched paragraph`);
    }
    return problems;
  } catch (error) {
    return [`threw ${error instanceof Error ? error.message : String(error)}`];
  } finally {
    for (const session of sessions) session.destroy();
  }
}

for (const [engine, yrs] of [['wasm', wasmYrs], ['native', nativeYrs as unknown as Yrs]] as const) {
  for (const kind of ['page', 'column']) {
    test(`${engine}: ${kind} breaks next to an edited paragraph boundary are saved as the session holds them`, async () => {
      const failures: string[] = [];
      let ran = 0;
      for (const [shape, edit, keepsL] of ROWS) {
        for (const mode of MODES) {
          ran += 1;
          for (const problem of await row(yrs, kind, shape, edit, mode, keepsL)) failures.push(`${shape} | ${edit} | ${mode}: ${problem}`);
        }
      }
      expect(ran).toBe(ROWS.length * MODES.length);
      expect(failures).toEqual([]);
    });
  }
}

/** Paragraphs whose break sits inside a non-run child, left untouched while another paragraph is edited. */
const NESTED: Record<string, string> = {
  'hyperlink starting with the break': `<w:hyperlink w:anchor="x"><w:r>${brk('page')}<w:t>A</w:t></w:r></w:hyperlink>`,
  'hyperlink with the break inside its text': `<w:hyperlink w:anchor="x"><w:r><w:t>A</w:t>${brk('page')}<w:t>C</w:t></w:r></w:hyperlink>`,
  'inline content control with the break inside its text': `<w:sdt><w:sdtContent><w:r><w:t>A</w:t>${brk('page')}<w:t>C</w:t></w:r></w:sdtContent></w:sdt>`,
  'inline content control starting with the break': `<w:sdt><w:sdtContent><w:r>${brk('page')}<w:t>A</w:t></w:r></w:sdtContent></w:sdt>`,
  'field result holding the break': `<w:r><w:t>A</w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r>${brk('page')}<w:t>1</w:t></w:r></w:fldSimple>`,
  'tracked insertion of the break': `<w:r><w:t>A</w:t></w:r><w:ins w:id="9" w:author="Ada" w:date="2026-01-01T00:00:00Z"><w:r>${brk('page')}</w:r></w:ins><w:r><w:t>C</w:t></w:r>`,
};

for (const [engine, yrs] of [['wasm', wasmYrs], ['native', nativeYrs as unknown as Yrs]] as const) {
  test(`${engine}: an untouched paragraph with a nested page break keeps it when another paragraph is edited`, async () => {
    const failures: string[] = [];
    for (const [name, inner] of Object.entries(NESTED)) {
      const source = `<w:p w14:paraId="72000001">${inner}</w:p>`;
      const bytes = await docx(L + source + N);
      const session = await yrs.createYrsSession();
      try {
        session.openDocx(bytes, true);
        session.insertText({ story: 'body', ...at(session, 2, 1) }, 'x');
        const live = blockBreaks(session, 'page');
        const { xml } = await savedXml(yrs, session);
        const written = (xml.match(/w:type="page"/g) ?? []).length;
        if (live !== 1 || written !== live) failures.push(`${name}: saved ${written} page breaks, live ${live}, the source 1`);
        if (!xml.includes(source)) failures.push(`${name}: rewrote the untouched paragraph`);
      } catch (error) {
        failures.push(`${name}: threw ${error instanceof Error ? error.message : String(error)}`);
      } finally {
        session.destroy();
      }
    }
    expect(failures).toEqual([]);
  });

  test(`${engine}: a text edit and a section break save beside a content control that starts with a page break`, async () => {
    const control = `<w:p w14:paraId="72000001">${NESTED['inline content control starting with the break']}</w:p>`;
    const bytes = await docx(L + control + N);
    const session = await yrs.createYrsSession();
    try {
      session.openDocx(bytes, true);
      session.insertText({ story: 'body', ...at(session, 2, 1) }, 'x');
      session.insertSectionBreakAfter(ids(session)[0]!, 'nextPage');
      const live = blockBreaks(session, 'page');
      const { xml } = await savedXml(yrs, session);
      expect([live, (xml.match(/w:type="page"/g) ?? []).length]).toEqual([1, 1]);
      expect(xml).toContain(control);
    } finally {
      session.destroy();
    }
  });
}
