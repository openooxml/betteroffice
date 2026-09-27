import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import {
  createYrsSession,
  type DocxEditRequest,
  type DocxEditResult,
  type DocxEditStep,
  type YrsSession,
} from './index';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const OFFICE = 'application/vnd.openxmlformats-officedocument';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  `xmlns:r="${REL}"`,
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:bofx="urn:fidelity"',
].join(' ');
const IMAGE_BYTES = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 1, 2, 3, 4]);
const CUSTOM_XML = '<?xml version="1.0" encoding="UTF-8"?><root xmlns="urn:custom"><value>kept</value></root>';
const CORE = '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Batch fixture</dc:title></cp:coreProperties>';
const RAW = '<bofx:block bofx:value="opaque"><bofx:child>hidden</bofx:child></bofx:block>';
const DRAWING = '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const paragraph = (id: string, content: string) => `<w:p w14:paraId="${id}">${content}</w:p>`;

const BODY = `${paragraph('00000001', `${run('Alpha')}<w:r><w:br/></w:r>${run('target words')}`)}${RAW}${paragraph('00000002', run('Second paragraph'))}<w:sdt><w:sdtPr><w:alias w:val="Clause"/><w:tag w:val="clause"/></w:sdtPr><w:sdtContent>${paragraph('00000003', run('Inside control'))}</w:sdtContent></w:sdt>${paragraph('00000004', `${run('Picture ')}${DRAWING}`)}${paragraph('00000005', run('Last'))}`;

function fixture(body = BODY): Uint8Array {
  const parts: PartsMap = new Map();
  const set = (name: string, content: string | Uint8Array) => parts.set(name, toBytes(content));
  set(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="${OFFICE}.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="${OFFICE}.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`
  );
  set(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`
  );
  set(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rIdHeader" Type="${REL}/header" Target="header1.xml"/><Relationship Id="rIdFooter" Type="${REL}/footer" Target="footer1.xml"/><Relationship Id="rIdImage" Type="${REL}/image" Target="media/image1.png"/><Relationship Id="rIdCustom" Type="${REL}/customXml" Target="../customXml/item1.xml"/></Relationships>`
  );
  set(
    'word/styles.xml',
    `<w:styles ${NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/></w:rPr></w:style></w:styles>`
  );
  set(
    'word/document.xml',
    `<w:document ${NS}><w:body>${body}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/></w:sectPr></w:body></w:document>`
  );
  set('word/header1.xml', `<w:hdr ${NS}>${paragraph('0000E001', run('Header text'))}</w:hdr>`);
  set('word/footer1.xml', `<w:ftr ${NS}>${paragraph('0000F001', run('Footer text'))}${RAW}</w:ftr>`);
  set('word/media/image1.png', IMAGE_BYTES);
  set('customXml/item1.xml', CUSTOM_XML);
  set('docProps/core.xml', CORE);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function part(bytes: ArrayBuffer | Uint8Array, name: string): Uint8Array {
  const entry = unzipContainer(new Uint8Array(bytes))[name];
  if (!(entry instanceof Uint8Array)) throw new Error(`missing part ${name}`);
  return entry;
}

const text = (bytes: Uint8Array) => new TextDecoder().decode(bytes);

/** Saved body order: each run text, and `RAW` for each opaque block. */
function blockOrder(xml: string): string[] {
  return [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<bofx:block /g)].map((match) => match[1] ?? 'RAW');
}

function texts(session: YrsSession, story = 'body', view: 'accepted' | 'original' = 'accepted') {
  const read = session.readParagraphs({ story, view });
  if (!read.ok) throw new Error(read.failure.message);
  return read.paragraphs.map((entry) => entry.text);
}

function applied(result: DocxEditResult): Extract<DocxEditResult, { ok: true }> {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result;
}

const search = (value: string, paraId: string, story = 'body') =>
  ({
    kind: 'search',
    text: value,
    within: { kind: 'paragraph', story, paraId },
    view: 'accepted',
  }) as const;

let nextClientId = 77100;

async function open(bytes = fixture()): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: nextClientId++ });
  session.openDocx(bytes, true);
  return session;
}

describe('YrsSession edit batches', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('reads versioned text with one U+FFFC per atom and finds reusable ranges', async () => {
    const session = await open();
    try {
      const read = session.readParagraphs({ paraIds: ['00000001', '00000004'], view: 'accepted' });
      expect(read).toMatchObject({
        ok: true,
        version: session.version(),
        paragraphs: [
          { story: 'body', paraId: '00000001', text: 'Alpha\uFFFCtarget words', atoms: [{ offset: 5, kind: 'lineBreak' }] },
          { paraId: '00000004', text: 'Picture \uFFFC', atoms: [{ offset: 8, kind: 'image' }] },
        ],
      });
      const found = session.findText({ text: 'target', within: { kind: 'story', story: 'body' }, view: 'accepted' });
      expect(found).toMatchObject({ ok: true, truncated: false, matches: [{ text: 'target' }] });
      if (!found.ok) throw new Error('search failed');
      const range = found.matches[0]!.range;
      expect([range.start.offset, range.end.offset]).toEqual([6, 12]);
      const result = applied(
        session.applyEdits({
          expectVersion: found.version,
          steps: [{ op: 'replaceText', target: { kind: 'range', ...range }, text: 'chosen' }],
        })
      );
      expect(result.applied).toBe(true);
      expect(texts(session)[0]).toBe('Alpha\uFFFCchosen words');
      expect(session.readParagraphs({ story: 'missing', view: 'accepted' })).toMatchObject({
        ok: false,
        failure: { code: 'missing-target' },
      });
    } finally {
      session.destroy();
    }
  });

  it('commits one notification per applied batch and none for refusals or no-ops', async () => {
    const session = await open();
    const replica = await createYrsSession({ clientId: nextClientId++ });
    try {
      replica.loadState(session.encodeState());
      const events: Array<{ update: Uint8Array; origin: string }> = [];
      const unsubscribe = session.onUpdate((update, origin) => events.push({ update, origin }));
      const version = session.version();
      const refused = session.applyEdits({
        expectVersion: version,
        steps: [
          { op: 'replaceText', target: search('Second', '00000002'), text: 'First' },
          { op: 'replaceText', target: search('absent', '00000005'), text: 'x' },
        ],
      });
      expect(refused).toMatchObject({ ok: false, version, failure: { code: 'missing-target', stepIndex: 1 } });
      const noOp = applied(
        session.applyEdits({
          expectVersion: version,
          steps: [{ op: 'replaceText', target: search('Second', '00000002'), text: 'Second' }],
        })
      );
      expect(noOp).toMatchObject({ applied: false, version, changedStories: [] });
      expect(events).toHaveLength(0);
      const result = applied(
        session.applyEdits({
          expectVersion: version,
          source: 'agent',
          steps: [
            { op: 'replaceText', target: search('Second', '00000002'), text: 'First' },
            {
              op: 'replaceText',
              target: { kind: 'paragraph', story: 'hf:rIdHeader', paraId: '0000E001' },
              text: 'New header',
            },
          ],
        })
      );
      expect(result).toMatchObject({ applied: true, source: 'agent', baseVersion: version, version: session.version() });
      expect(result.changedStories).toEqual(['body', 'hf:rIdHeader']);
      expect(events.map((event) => event.origin)).toEqual(['local']);
      replica.applyUpdate(events[0]!.update);
      expect(texts(replica)).toEqual(texts(session));
      expect(texts(replica, 'hf:rIdHeader')).toEqual(['New header']);
      unsubscribe();
    } finally {
      replica.destroy();
      session.destroy();
    }
  });

  it('throws on malformed requests and refuses stale versions as data', async () => {
    const session = await open();
    const peer = await createYrsSession({ clientId: nextClientId++ });
    try {
      expect(() => session.applyEdits({ steps: [] } as unknown as DocxEditRequest)).toThrow();
      expect(() =>
        session.validateEdits({
          expectVersion: session.version(),
          steps: [{ op: 'moveText' } as unknown as DocxEditStep],
        })
      ).toThrow();
      expect(() =>
        session.applyEdits({
          expectVersion: session.version(),
          steps: [
            {
              op: 'deleteText',
              target: { kind: 'range', story: 'body', start: { paraId: 'p', offset: -1 }, end: { paraId: 'p', offset: 0 }, view: 'accepted' },
            },
          ],
        })
      ).toThrow();
      const request: DocxEditRequest = {
        expectVersion: session.version(),
        steps: [{ op: 'insertText', target: search('Last', '00000005'), at: 'end', text: '!' }],
      };
      expect(session.validateEdits(request)).toMatchObject({ ok: true, wouldApply: true });
      peer.loadState(session.encodeState());
      peer.insertText({ story: 'body', paraId: '00000005', offset: 0 }, 'Very ');
      session.applyUpdate(peer.encodeState());
      expect(session.validateEdits(request)).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      expect(session.applyEdits(request)).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      expect(texts(session).at(-1)).toBe('Very Last');
    } finally {
      peer.destroy();
      session.destroy();
    }
  });

  it('rotates the version when the same session reopens its document', async () => {
    const bytes = fixture();
    const session = await createYrsSession({ clientId: nextClientId++ });
    const other = await open(bytes);
    try {
      session.openDocx(bytes, false);
      const first = session.version();
      session.openDocx(bytes, false);
      expect(session.version()).not.toBe(first);
      session.loadState(other.encodeState());
      const joined = session.version();
      const result = applied(
        session.applyEdits({
          expectVersion: joined,
          steps: [
            {
              op: 'insertParagraphs',
              target: { story: 'body', paraId: '00000005' },
              at: 'end',
              paragraphs: [{ text: 'Joined', styleId: 'Heading1' }],
            },
          ],
        })
      );
      expect(result.receipts[0]!.newParagraphs).toHaveLength(1);
      expect(
        other.applyEdits({ expectVersion: joined, steps: [] })
      ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    } finally {
      other.destroy();
      session.destroy();
    }
  });

  it('keeps each applied batch one undo step, or none when history is none', async () => {
    const session = await open();
    try {
      session.beginUndoCapture();
      session.insertText({ story: 'body', paraId: '00000005', offset: 4 }, '.');
      session.addUndoBoundary();
      applied(
        session.applyEdits({
          expectVersion: session.version(),
          steps: [
            { op: 'replaceText', target: search('Second', '00000002'), text: 'Next' },
            { op: 'deleteParagraphs', story: 'body', firstParaId: '00000004', lastParaId: '00000004' },
          ],
        })
      );
      expect(texts(session)).toEqual(['Alpha\uFFFCtarget words', 'Next paragraph', 'Last.']);
      expect(session.undo()).toBe(true);
      expect(texts(session)).toEqual(['Alpha\uFFFCtarget words', 'Second paragraph', 'Picture \uFFFC', 'Last.']);
      expect(session.canRedo()).toBe(true);
      applied(
        session.applyEdits({
          expectVersion: session.version(),
          history: 'none',
          steps: [{ op: 'insertText', target: search('Alpha', '00000001'), at: 'start', text: '>' }],
        })
      );
      expect(session.canRedo()).toBe(true);
      expect(session.undo()).toBe(true);
      expect(texts(session)[0]).toBe('>Alpha\uFFFCtarget words');
      expect(texts(session).at(-1)).toBe('Last');
    } finally {
      session.destroy();
    }
  });

  it('routes the legacy helpers through the shared resolver after a hard break', async () => {
    const session = await open();
    try {
      const formatted = session.formatTextTarget(search('target', '00000001'), { bold: true });
      expect(formatted).toMatchObject({ ok: true, version: session.version() });
      const bold = session
        .storySegments('body')
        .filter((segment) => segment.kind === 'text' && segment.attributes.bold === true)
        .map((segment) => (segment.kind === 'text' ? segment.text : ''));
      expect(bold).toEqual(['target']);
      expect(
        session.commentTextTarget(search('words', '00000001'), { id: '41', author: 'Ann', date: '2026-09-24T00:00:00Z', body: null })
      ).toMatchObject({ ok: true });
      const [anchor] = session.resolveComment('41');
      const span = session.locateParagraph('body', '00000001');
      expect([anchor!.start - span.start, anchor!.end - span.start]).toEqual([13, 18]);
      expect(session.formatTextTarget(search('absent', '00000001'), { bold: true })).toMatchObject({
        ok: false,
        failure: { code: 'missing-target' },
      });
      expect(
        session.selectionText({
          story: 'body',
          start: { paraId: '00000001', offset: 2 },
          end: { paraId: '00000001', offset: 12 },
        })
      ).toEqual({
        paraId: '00000001',
        selectedText: 'pha\uFFFCtarget',
        paragraphText: 'Alpha\uFFFCtarget words',
        before: 'Al',
        after: ' words',
      });
    } finally {
      session.destroy();
    }
  });

  it('saves and reopens batch edits without losing unrelated package content', async () => {
    const bytes = fixture();
    const session = await open(bytes);
    const reopened = await createYrsSession({ clientId: nextClientId++ });
    try {
      const base = session.materializeDocx();
      if (!base) throw new Error('the opened package must materialize');
      yrsToDocument(session, base);
      const read = session.readParagraphs({ view: 'accepted' });
      if (!read.ok) throw new Error(read.failure.message);
      const result = applied(
        session.applyEdits({
          expectVersion: read.version,
          steps: [
            {
              op: 'replaceText',
              target: search('target', '00000001'),
              text: 'chosen',
              expect: { text: 'target' },
            },
            {
              op: 'insertParagraphs',
              target: { story: 'body', paraId: '00000005' },
              at: 'end',
              paragraphs: [{ text: 'Appended heading', styleId: 'Heading1' }, { text: 'Appended body' }],
            },
            {
              op: 'replaceText',
              target: { kind: 'paragraph', story: 'hf:rIdHeader', paraId: '0000E001' },
              text: 'Edited header',
            },
            {
              op: 'replaceText',
              target: { kind: 'paragraph', story: 'body:sdt0', paraId: '00000003' },
              text: 'Filled control',
            },
            {
              op: 'replaceText',
              target: search('Last', '00000005'),
              text: 'Final',
              suggest: { author: 'Reviewer', date: '2026-09-24T12:00:00Z' },
            },
          ],
        })
      );
      expect(result.receipts[4]!.revisionIds).toHaveLength(1);
      expect(result.changedStories).toEqual(['body', 'body:sdt0', 'hf:rIdHeader']);
      const saved = await repackDocx(yrsToDocument(session, base));
      for (const name of ['word/media/image1.png', 'customXml/item1.xml']) {
        expect(part(saved, name)).toEqual(part(bytes, name));
      }
      expect(text(part(saved, 'docProps/core.xml'))).toContain('<dc:title>Batch fixture</dc:title>');
      const document = text(part(saved, 'word/document.xml'));
      expect(document).toContain(RAW);
      expect(document).toContain('<w:tag w:val="clause"/>');
      expect(document).toContain('r:embed="rIdImage"');
      expect(document).toContain('Filled control');
      expect(document).toMatch(/<w:ins [^>]*w:author="Reviewer"/);
      expect(document).toMatch(/<w:del [^>]*w:author="Reviewer"/);
      expect(text(part(saved, 'word/footer1.xml'))).toContain(RAW);
      expect(text(part(saved, 'word/header1.xml'))).toContain('Edited header');
      expect(text(part(saved, 'word/_rels/document.xml.rels'))).toContain('Target="../customXml/item1.xml"');

      reopened.openDocx(new Uint8Array(saved), true);
      expect(texts(reopened)).toEqual([
        'Alpha\uFFFCchosen words',
        'Second paragraph',
        'Picture \uFFFC',
        'Final',
        'Appended heading',
        'Appended body',
      ]);
      expect(texts(reopened, 'body', 'original')[3]).toBe('Last');
      expect(texts(reopened, 'body:sdt0')).toEqual(['Filled control']);
      expect(texts(reopened, 'hf:rIdHeader')).toEqual(['Edited header']);
      const heading = reopened.readParagraphs({ view: 'accepted' });
      if (!heading.ok) throw new Error(heading.failure.message);
      expect(heading.paragraphs[4]!.styleId).toBe('Heading1');
      expect(
        reopened.applyEdits({
          expectVersion: result.version,
          steps: [{ op: 'replaceText', target: search('Last', '00000005'), text: 'x' }],
        })
      ).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
    } finally {
      reopened.destroy();
      session.destroy();
    }
  });

  it('restores opaque blocks before the live paragraph that followed them', async () => {
    const session = await open(
      fixture(`${paragraph('00000001', run('A'))}${RAW}${paragraph('00000002', run('B'))}${paragraph('00000003', run('C'))}`)
    );
    try {
      const base = session.materializeDocx();
      if (!base) throw new Error('the opened package must materialize');
      const order = async () =>
        blockOrder(text(part(await repackDocx(yrsToDocument(session, base)), 'word/document.xml')));
      const insertBeforeC = () =>
        session.applyEdits({
          expectVersion: session.version(),
          steps: [{ op: 'insertParagraphs', target: { story: 'body', paraId: '00000003' }, at: 'start', paragraphs: [{ text: 'X' }] }],
        });
      const deleteB = () =>
        session.applyEdits({
          expectVersion: session.version(),
          steps: [{ op: 'deleteParagraphs', story: 'body', firstParaId: '00000002', lastParaId: '00000002' }],
        });
      session.beginUndoCapture();
      applied(insertBeforeC());
      expect(await order()).toEqual(['A', 'RAW', 'B', 'X', 'C']);
      expect(deleteB()).toMatchObject({ ok: false, failure: { code: 'unsupported' } });
      expect(session.undo()).toBe(true);
      applied(deleteB());
      expect(await order()).toEqual(['A', 'RAW', 'C']);
      expect(insertBeforeC()).toMatchObject({ ok: false, failure: { code: 'unsupported' } });
      expect(session.undo()).toBe(true);
      expect(await order()).toEqual(['A', 'RAW', 'B', 'C']);
    } finally {
      session.destroy();
    }
  });

  it('refuses steps whose save would move an opaque block past a table', async () => {
    const table = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${paragraph('0000C001', run('cell'))}</w:tc></w:tr></w:tbl>`;
    const displaced = await open(
      fixture(`${paragraph('00000001', run('A'))}${RAW}${table}${paragraph('00000002', run('B'))}`)
    );
    const inPlace = await open(
      fixture(`${paragraph('00000001', run('A'))}${table}${RAW}${paragraph('00000002', run('B'))}`)
    );
    const reopened = await createYrsSession({ clientId: nextClientId++ });
    try {
      const insertAfterA = (session: YrsSession) =>
        session.applyEdits({
          expectVersion: session.version(),
          steps: [{ op: 'insertParagraphs', target: { story: 'body', paraId: '00000001' }, at: 'end', paragraphs: [{ text: 'X' }] }],
        });
      const editCell = (session: YrsSession) =>
        session.applyEdits({
          expectVersion: session.version(),
          steps: [{ op: 'replaceText', target: { kind: 'paragraph', story: 'body:t0:r0c0', paraId: '0000C001' }, text: 'edited' }],
        });
      expect(insertAfterA(displaced)).toMatchObject({ ok: false, failure: { code: 'unsupported' } });
      expect(editCell(displaced)).toMatchObject({ ok: false, failure: { code: 'unsupported' } });
      const base = inPlace.materializeDocx();
      if (!base) throw new Error('the opened package must materialize');
      applied(insertAfterA(inPlace));
      applied(editCell(inPlace));
      const saved = await repackDocx(yrsToDocument(inPlace, base));
      expect(blockOrder(text(part(saved, 'word/document.xml')))).toEqual(['A', 'X', 'edited', 'RAW', 'B']);
      reopened.openDocx(new Uint8Array(saved), true);
      expect(texts(reopened)).toEqual(['A', 'X', 'B']);
      expect(texts(reopened, 'body:t0:r0c0')).toEqual(['edited']);
      expect(editCell(reopened)).toMatchObject({ ok: true, applied: false });
    } finally {
      reopened.destroy();
      inPlace.destroy();
      displaced.destroy();
    }
  });

  it('refuses edits that would drop tracked run formatting and keeps it through other edits', async () => {
    const changed =
      '<w:r><w:rPr><w:b/><w:rPrChange w:id="7" w:author="Ann" w:date="2026-09-24T12:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>bold</w:t></w:r>';
    const session = await open(
      fixture(`${paragraph('00000001', `${run('plain ')}${changed}`)}${paragraph('00000002', run('Other'))}`)
    );
    const reopened = await createYrsSession({ clientId: nextClientId++ });
    try {
      const base = session.materializeDocx();
      if (!base) throw new Error('the opened package must materialize');
      const touch = (target: YrsSession) =>
        target.applyEdits({
          expectVersion: target.version(),
          steps: [{ op: 'insertText', target: search('plain', '00000001'), at: 'end', text: '!' }],
        });
      expect(touch(session)).toMatchObject({ ok: false, failure: { code: 'tracked-revision-conflict' } });
      applied(
        session.applyEdits({
          expectVersion: session.version(),
          steps: [{ op: 'replaceText', target: search('Other', '00000002'), text: 'Changed' }],
        })
      );
      const saved = await repackDocx(yrsToDocument(session, base));
      expect(text(part(saved, 'word/document.xml'))).toMatch(/<w:rPrChange [^>]*w:author="Ann"/);
      reopened.openDocx(new Uint8Array(saved), true);
      expect(texts(reopened)).toEqual(['plain bold', 'Changed']);
      expect(touch(reopened)).toMatchObject({ ok: false, failure: { code: 'tracked-revision-conflict' } });
    } finally {
      reopened.destroy();
      session.destroy();
    }
  });
});
