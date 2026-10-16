import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { createYrsSession } from './index';
import { preloadEditWasm } from '../wasm/edit';
import type { AnchorReader } from './proposalGeometry';
import { resolveReservations, validReservationsPayload, type DocxReservationsPayload } from './reservations';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(
  import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'
)))));

async function sessionWithParagraphs(count: number, before = '') {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const body = Array.from({ length: count }, (_, index) => `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}"><w:r><w:t>Paragraph ${index}</w:t></w:r></w:p>`).join('');
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${before}${body}<w:sectPr/></w:body></w:document>`);
  const session = await createYrsSession();
  session.openDocx(await zip.generateAsync({ type: 'uint8array' }), true);
  return session;
}

function countingReader(reader: AnchorReader) {
  const counts = { segments: 0, spans: 0, builds: 0 };
  const wrapped: AnchorReader = {
    resolveParagraphAnchor: (anchor) => reader.resolveParagraphAnchor(anchor),
    hasStory: (story) => reader.hasStory(story),
    findText: (...args) => reader.findText(...args),
    listRevisions: () => reader.listRevisions(),
    paragraphSpans: (story) => { counts.spans += 1; return reader.paragraphSpans(story); },
    storySegments: (story) => {
      counts.segments += 1;
      const segments = reader.storySegments(story);
      return new Proxy(segments, {
        get(target, key, receiver) {
          if (key === Symbol.iterator) return () => { counts.builds += 1; return target[Symbol.iterator](); };
          return Reflect.get(target, key, receiver);
        },
      });
    },
  };
  return { reader: wrapped, counts };
}

function payloadFor(session: Awaited<ReturnType<typeof sessionWithParagraphs>>, count: number): DocxReservationsPayload {
  const version = session.version();
  const paragraphs = session.paragraphs('body');
  return { version, sets: [{ owner: 'a', items: paragraphs.slice(0, count).map((paragraph, index) => ({
    id: `r${index}`, kind: 'hide', anchor: { kind: 'range', version, range: {
      story: 'body', view: 'accepted', start: { paraId: paragraph.paraId, offset: 1 },
      end: { paraId: paragraphs[Math.min(index + 1, paragraphs.length - 1)]!.paraId, offset: 3 },
    } },
  })) }] };
}

test('reservation carried sets retain only their previous owner entries', async () => {
  const session = await sessionWithParagraphs(3);
  try {
    const original = payloadFor(session, 2);
    const previous = resolveReservations(session, { ...original, sets: [
      ...original.sets, { owner: 'removed', items: original.sets[0]!.items },
    ] });
    session.insertText({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 0 }, 'X');
    const carried = { version: session.version(), sets: [{ ...original.sets[0]!, version: original.version }] };
    expect(validReservationsPayload(carried)).toBe(true);
    const result = resolveReservations(session, carried, session.version(), previous);
    const retained = Object.fromEntries(Object.entries(previous).filter(([, entry]) => entry.owner === 'a'));
    expect(result).toEqual(retained);
    expect(validReservationsPayload({ ...carried, sets: [{ ...carried.sets[0], version: '' }] })).toBe(false);
    for (const [key, entry] of Object.entries(result)) expect(entry).toBe(previous[key]);
    expect(resolveReservations(session, carried, session.version(), null)).toEqual({});
  } finally { session.destroy(); }
});

test('reservation range resolution builds story views and spans once for a thousand anchors', async () => {
  const session = await sessionWithParagraphs(1000);
  try {
    const counted = countingReader(session);
    const resolved = resolveReservations(counted.reader, payloadFor(session, 1000));
    expect(Object.keys(resolved).length).toBe(1999);
    expect(counted.counts.segments).toBeLessThanOrEqual(1);
    expect(counted.counts.builds).toBeLessThanOrEqual(1);
    expect(counted.counts.spans).toBeLessThanOrEqual(1);
    const small = payloadFor(session, 3);
    const unmemoized = Object.assign({}, ...small.sets[0]!.items.map((item) => resolveReservations(
      countingReader(session).reader, { ...small, sets: [{ owner: 'a', items: [item] }] }
    )));
    expect(resolveReservations(countingReader(session).reader, small)).toEqual(unmemoized);
  } finally { session.destroy(); }
});

test('a paragraph opening after a table resolves from its first inline unit, as layout counts it', async () => {
  const table = '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const session = await sessionWithParagraphs(1, table);
  try {
    const version = session.version();
    const { paraId, text } = session.paragraphs('body')[0]!;
    const paragraph = { kind: 'session' as const, sessionId: session.paragraphIdentities().sessionId, story: 'body', paraId };
    const at = (offset: number) => ({ paraId, offset });
    const resolved = resolveReservations(session, { version, sets: [{ owner: 'a', items: [
      { id: 'cut', kind: 'hide', anchor: { kind: 'range', version, range: { story: 'body', view: 'accepted', start: at(1), end: at(3) } } },
      { id: 'gone', kind: 'hide', anchor: { kind: 'paragraph', paragraph } },
    ] }] });
    expect(resolved).toEqual({
      [JSON.stringify(['a', 'cut'])]: { owner: 'a', id: 'cut', story: 'body', paraId, start: 1, end: 3, kind: 'hide' },
      [JSON.stringify(['a', 'gone'])]: { owner: 'a', id: 'gone', story: 'body', paraId, start: 0, end: text.length + 1, kind: 'hide' },
    });
  } finally { session.destroy(); }
});
