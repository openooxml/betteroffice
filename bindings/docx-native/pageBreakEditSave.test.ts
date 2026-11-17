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
  'select from the break paragraph start into the next paragraph': (s, a) => { s.deleteRange({ story: 'body', start: at(s, 1, 0), end: at(s, 2, 1) }, a); },
  'select from the break paragraph end into the next paragraph': (s, a) => { s.deleteRange({ story: 'body', start: at(s, 1, length(s, 1)), end: at(s, 2, 1) }, a); },
  'remove the break through the raw story API': (s) => { s.applyRawOps('body', [{ op: 'delete', index: breakIndex(s), len: 1 }]); },
  'split at the end of the break paragraph': (s, a) => { s.splitParagraph({ story: 'body', ...at(s, 1, length(s, 1)) }, a); },
  'merge with the next paragraph': (s, a) => { s.mergeParagraphs('body', ids(s)[1]!, a); },
  'type before the break': (s, a) => { s.insertText({ story: 'body', ...at(s, 1, 1) }, 'x', a); },
  'Enter at the end of the previous paragraph': (s, a) => { s.splitParagraph({ story: 'body', ...at(s, 0, 1) }, a); },
  'join the next paragraph': (s, a) => { s.deleteRange({ story: 'body', start: at(s, 1, length(s, 1)), end: at(s, 2, 0) }, a); },
};

function breakIndex(session: Session): number {
  let index = 0;
  for (const segment of session.storySegments('body')) {
    if (segment.kind === 'embed' && (segment.embedKind === 'pageBreak' || segment.embedKind === 'columnBreak')) return index;
    index += segment.kind === 'text' ? segment.text.length : 1;
  }
  throw new Error('no break in the story');
}

const MODES = ['plain', 'undo and redo', 'tracked', 'remote peer'] as const;
type Mode = typeof MODES[number];

/**
 * [shape, edit, the paragraph `L` stays as its source bytes]. Deleting a break that follows
 * text in its paragraph keeps the break on save, as before: the saved paragraph cannot tell a
 * removed break from one the editing model did not place.
 */
const ROWS: Array<[string, string, boolean]> = [
  ...['middle', 'middle in three runs', 'before a bare section mark'].flatMap((shape): Array<[string, string, boolean]> => [
    [shape, 'split before the break', true],
    [shape, 'join with the previous paragraph', false],
    [shape, 'merge the previous paragraph', false],
    [shape, 'remove the text past the break', true],
    [shape, 'split at the end of the break paragraph', true],
    [shape, 'delete the break paragraph', false],
    [shape, 'select from the break paragraph start into the next paragraph', true],
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
  ['end', 'select from the break paragraph start into the next paragraph', true],
  ['end', 'select from the break paragraph end into the next paragraph', true],
  ['own paragraph', 'select from the break paragraph start into the next paragraph', true],
  ['own paragraph', 'select from the break paragraph end into the next paragraph', true],
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

/**
 * The text that opens each page, read from the story: a page break embed starts the next page,
 * and a pilcrow's pageBreakBeforeRun starts a page at its own paragraph's first text.
 */
function pageStarts(session: Session): string[] {
  const starts: string[] = [''];
  let pending = false;
  let texts: string[] = [];
  const flush = (leading: boolean) => {
    if (leading) pending = true;
    for (const value of texts) {
      if (pending) { starts.push(''); pending = false; }
      if (!starts[starts.length - 1]) starts[starts.length - 1] = value.slice(0, 3);
    }
    texts = [];
  };
  for (const segment of session.storySegments('body')) {
    if (segment.kind === 'embed' && segment.embedKind === 'pageBreak') { flush(false); pending = true; continue; }
    if (segment.kind === 'pilcrow') { flush(segment.properties.pageBreakBeforeRun === true); continue; }
    if (segment.kind === 'text' && segment.text.trim()) texts.push(segment.text.trim());
  }
  flush(false);
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
      const next = save.xml.match(/<w:p w14:paraId="72000003"[^>]*>[\s\S]*?<\/w:p>/)?.[0];
      if (next?.includes(`<w:r>${brk(kind)}</w:r>`) && next.replace(/<[^>]*>/g, '') === 'N' && next.replaceAll(`<w:r>${brk(kind)}</w:r>`, '') !== N) {
        problems.push(`${who} wrote the paragraph after the break as ${next}, not its source with only break runs added`);
      }
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
      for (const [edited, index] of [['the previous paragraph', 0]] as const) {
        const source = `<w:p w14:paraId="72000001">${inner}</w:p>`;
        const bytes = await docx(L + source + N);
        const session = await yrs.createYrsSession();
        try {
          session.openDocx(bytes, true);
          session.insertText({ story: 'body', ...at(session, index, 1) }, 'x');
          const live = blockBreaks(session, 'page');
          const { xml } = await savedXml(yrs, session);
          const written = (xml.match(/w:type="page"/g) ?? []).length;
          if (live !== 1 || written !== live) failures.push(`${name}, edit in ${edited}: saved ${written} page breaks, live ${live}, the source 1`);
          if (!xml.includes(source)) failures.push(`${name}, edit in ${edited}: rewrote the untouched paragraph`);
          if (index === 0 && !xml.includes(N)) failures.push(`${name}, edit in ${edited}: rewrote the paragraph after the break`);
        } catch (error) {
          failures.push(`${name}, edit in ${edited}: threw ${error instanceof Error ? error.message : String(error)}`);
        } finally {
          session.destroy();
        }
      }
    }
    expect(failures).toEqual([]);
  });

  test(`${engine}: untouched paragraphs around a table keep their page break when another paragraph is edited`, async () => {
    const broken = `<w:p w14:paraId="72000001"><w:r><w:t>A</w:t>${brk('page')}</w:r></w:p>`;
    const table = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p w14:paraId="72000011">${text('M')}</w:p></w:tc></w:tr></w:tbl>`;
    const session = await yrs.createYrsSession();
    try {
      session.openDocx(await docx(L + broken + table + N), true);
      session.insertText({ story: 'body', ...at(session, 0, 1) }, 'x');
      const { xml } = await savedXml(yrs, session);
      expect([blockBreaks(session, 'page'), (xml.match(/w:type="page"/g) ?? []).length]).toEqual([1, 1]);
      expect(xml).toContain(broken + table + N);
    } finally {
      session.destroy();
    }
  });

  test(`${engine}: a page break beside a hyperlink is written once`, async () => {
    const failures: string[] = [];
    const link = '<w:hyperlink w:anchor="x"><w:r><w:t>h</w:t></w:r></w:hyperlink>';
    const placements: Record<string, (p: string) => string> = {
      'before the next paragraph': (p) => L + p + N,
      'last in the body': (p) => L + p,
      'last in a table cell': (p) => L + `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p}</w:tc></w:tr></w:tbl>` + N,
      'before a table': (p) => L + p + `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p w14:paraId="72000011">${text('M')}</w:p></w:tc></w:tr></w:tbl>` + N,
    };
    for (const [where, place] of Object.entries(placements)) {
      for (const [shape, inner] of [['end', `${link}<w:r><w:t>AB</w:t>${brk('page')}</w:r>`], ['middle', `${link}<w:r><w:t>A</w:t>${brk('page')}<w:t>B</w:t></w:r>`]] as const) {
        const broken = `<w:p w14:paraId="72000001">${inner}</w:p>`;
        const session = await yrs.createYrsSession();
        try {
          session.openDocx(await docx(place(broken)), true);
          session.insertText({ story: 'body', ...at(session, 0, 1) }, 'x');
          const { xml } = await savedXml(yrs, session);
          const written = (xml.match(/w:type="page"/g) ?? []).length;
          if (written !== 1) failures.push(`${where}, ${shape}, edit elsewhere: saved ${written} page breaks`);
          if (!xml.includes(broken)) failures.push(`${where}, ${shape}, edit elsewhere: rewrote the paragraph with the break`);
        } catch (error) {
          failures.push(`${where}, ${shape}: threw ${error instanceof Error ? error.message : String(error)}`);
        } finally {
          session.destroy();
        }
      }
    }
    expect(failures).toEqual([]);
  });

  test(`${engine}: page starts read a break the same on either side of its boundary and tell a moved one apart`, async () => {
    const P = (inner: string) => `<w:p w14:paraId="72000001">${inner}</w:p>`;
    const M = (inner: string) => `<w:p w14:paraId="72000004">${inner}</w:p>`;
    const N2 = (inner: string) => `<w:p w14:paraId="72000003">${inner}</w:p>`;
    const bodies = {
      trailing: L + P(`<w:r><w:t>A</w:t>${brk('page')}</w:r>`) + N + M(text('M')),
      leading: L + P(text('A')) + N2(`<w:r>${brk('page')}<w:t>N</w:t></w:r>`) + M(text('M')),
      moved: L + P(text('A')) + N + M(`<w:r>${brk('page')}<w:t>M</w:t></w:r>`),
    };
    const starts: Record<string, string[]> = {};
    for (const [name, body] of Object.entries(bodies)) {
      const session = await yrs.createYrsSession();
      try {
        session.openDocx(await docx(body), true);
        starts[name] = pageStarts(session);
      } finally {
        session.destroy();
      }
    }
    expect(starts.trailing).toEqual(['L', 'N']);
    expect(starts.leading).toEqual(starts.trailing);
    expect(starts.moved).toEqual(['L', 'M']);
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
