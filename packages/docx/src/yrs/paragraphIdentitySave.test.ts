import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';

import objects from '../../../../crates/docx-edit/tests/fixtures/behind_objects.json';
import { rezipPartsToArrayBuffer, type PartsMap } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import {
  createYrsSession,
  saveYrsDocx,
  type DocxParagraphAnchor,
  type DocxParagraphRef,
  type DocxPersistedParagraphAnchor,
  type DocxSourceStory,
  type YrsSession,
} from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FIXTURE = resolve(
  import.meta.dir,
  '../../../../crates/docx-edit/tests/fixtures/paragraph-identities'
);
const BODY: DocxSourceStory = { partUri: '/word/document.xml', kind: 'body' };
const DOCUMENT = 'word/document.xml';
const FOOTNOTES = 'word/footnotes.xml';
const TAIL = '<w:p w14:paraId="0A0B0C0D" w:rsidR="00A1B2C3"><w:r><w:t>Tail</w:t></w:r></w:p>';
const DROPDOWN =
  '<w:sdt><w:sdtPr><w:id w:val="42"/><w:dropDownList>' +
  '<w:listItem w:displayText="Alpha" w:value="a"/><w:listItem w:displayText="Beta" w:value="b"/>' +
  '</w:dropDownList></w:sdtPr><w:sdtContent>' +
  '<w:p w14:paraId="3A3B3C3D"><w:r><w:t>Alpha</w:t></w:r></w:p>' +
  '<w:p w14:paraId="3A3B3C3E"><w:r><w:t>Second</w:t></w:r></w:p>' +
  '</w:sdtContent></w:sdt>';

function partNames(dir: string, prefix = ''): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) =>
    entry.isDirectory()
      ? partNames(join(dir, entry.name), `${prefix}${entry.name}/`)
      : [`${prefix}${entry.name}`]
  );
}

function fixtureParts(): Map<string, string> {
  return new Map(
    partNames(FIXTURE)
      .sort()
      .map((name) => [name, readFileSync(join(FIXTURE, name), 'utf8')])
  );
}

/** The shared identity fixture, with `[part, from, to]` rewrites applied. */
function fixture(...edits: Array<[string, string, string]>): Uint8Array {
  const parts = fixtureParts();
  for (const [name, from, to] of edits) {
    const xml = parts.get(name)!;
    expect(xml).toContain(from);
    parts.set(name, xml.replace(from, to));
  }
  const pkg: PartsMap = new Map(
    [...parts].map(([name, xml]) => [name, new TextEncoder().encode(xml)])
  );
  return new Uint8Array(rezipPartsToArrayBuffer(pkg));
}

function parts(bytes: Uint8Array): Record<string, Uint8Array> {
  return unzipContainer(bytes);
}

function text(bytes: Uint8Array, name: string): string {
  return new TextDecoder().decode(parts(bytes)[name]);
}

const STRICT_PARAGRAPHS = `import sys,io,zipfile,json,xml.etree.ElementTree as E
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14='{http://schemas.microsoft.com/office/word/2010/wordml}paraId'
z=zipfile.ZipFile(io.BytesIO(sys.stdin.buffer.read()))
result={}
for name in sys.argv[1:]:
 root=E.fromstring(z.read(name))
 found=[]
 for p in root.iter(W+'p'):
  for key in p.attrib:
   assert 'paraId' not in key or key==W14,(name,key)
  found.append([p.attrib.get(W14),''.join(t.text or '' for t in p.iter(W+'t'))])
 result[name]=found
print(json.dumps(result))`;

/**
 * Each `w:p` of the parts as `[w14:paraId, text]` in document order, read by
 * an independent strict XML parser that resolves attribute namespaces.
 */
function savedParagraphs(
  bytes: Uint8Array,
  ...names: string[]
): Record<string, Array<[string | null, string]>> {
  const result = spawnSync('python3', ['-c', STRICT_PARAGRAPHS, ...names], {
    input: Buffer.from(bytes),
    encoding: 'utf8',
  });
  expect(result.stderr).toBe('');
  expect(result.status).toBe(0);
  return JSON.parse(result.stdout) as Record<string, Array<[string | null, string]>>;
}

function idOf(paragraphs: Array<[string | null, string]>, text: string): string | null {
  const found = paragraphs.filter(([, value]) => value === text);
  expect(found).toHaveLength(1);
  return found[0]![0];
}

function expectAllocated(id: string | null): string {
  expect(id).toMatch(/^[0-9A-F]{8}$/);
  const value = Number.parseInt(id!, 16);
  expect(value).toBeGreaterThan(0);
  expect(value).toBeLessThanOrEqual(0x7ffffffe);
  return id!;
}

const sessions: YrsSession[] = [];

async function session(clientId: number): Promise<YrsSession> {
  const created = await createYrsSession({ clientId });
  sessions.push(created);
  return created;
}

async function open(bytes: Uint8Array, clientId: number): Promise<YrsSession> {
  const opened = await session(clientId);
  opened.openDocx(bytes, true);
  return opened;
}

function found(session: YrsSession, anchor: DocxParagraphAnchor): DocxParagraphRef {
  const result = session.resolveParagraphAnchor(anchor);
  if (result.status !== 'found') throw new Error(`unresolved: ${JSON.stringify(result)}`);
  return result.anchor;
}

function resolvedText(session: YrsSession, anchor: DocxParagraphAnchor): string | undefined {
  const paragraph = found(session, anchor);
  if (paragraph.kind !== 'session') throw new Error('not a session paragraph');
  return session.paragraphs(paragraph.story).find((entry) => entry.paraId === paragraph.paraId)
    ?.text;
}

function key(session: YrsSession, story: string, text: string): string {
  return session.paragraphs(story).find((paragraph) => paragraph.text === text)!.paraId;
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

afterEach(() => {
  for (const created of sessions.splice(0)) created.destroy();
});

describe('paragraph identities across saves', () => {
  it('resolves a session-authored paragraph after each save and reopen', async () => {
    const opened = await open(fixture(), 7);
    const split = opened.splitParagraph({ story: 'body', paraId: '1A2B3C4D', offset: 2 });
    const saved = await saveYrsDocx(opened);
    const authored = saved.paragraphs.find(
      (paragraph) => paragraph.session.paraId === split.secondParaId
    )!.persisted;
    expect(authored.story).toEqual(BODY);
    const body = savedParagraphs(saved.bytes, DOCUMENT)[DOCUMENT]!;
    expect(idOf(body, 'lid')).toBe(expectAllocated(authored.paraId));
    expect(idOf(body, 'Va')).toBe('1A2B3C4D');
    expect(text(saved.bytes, DOCUMENT)).not.toContain(split.secondParaId);

    const reopened = await open(saved.bytes, 8);
    expect(resolvedText(reopened, authored)).toBe('lid');
    const again = await saveYrsDocx(reopened);
    expect(again.paragraphs.map((paragraph) => paragraph.persisted)).toContainEqual(authored);
    const third = await open(again.bytes, 9);
    expect(resolvedText(third, authored)).toBe('lid');
    const repeated = third.resolveParagraphAnchor({ kind: 'persisted', story: BODY, paraId: '1A2B3C4D' });
    expect(repeated.status).toBe('ambiguous');
    expect(
      repeated.status === 'ambiguous' &&
        repeated.candidates.map((candidate) =>
          candidate.kind === 'session'
            ? third.paragraphs('body').find((entry) => entry.paraId === candidate.paraId)?.text
            : undefined
        )
    ).toEqual(['Va', 'Duplicate']);
  });

  it('keeps first halves and gives second halves new IDs at every split position', async () => {
    const opened = await open(fixture(), 7);
    const atStart = opened.splitParagraph({ story: 'body', paraId: '1A2B3C4D', offset: 0 });
    const atEnd = opened.splitParagraph({ story: 'body', paraId: '0000abcd', offset: 5 });
    opened.insertText({ story: 'body', paraId: atEnd.secondParaId, offset: 0 }, 'After');
    const saved = await saveYrsDocx(opened);
    const body = savedParagraphs(saved.bytes, DOCUMENT)[DOCUMENT]!;
    expect(body[0]).toEqual(['1A2B3C4D', '']);
    const valid = expectAllocated(idOf(body, 'Valid'));
    expect(idOf(body, 'Lower')).toBe('0000abcd');
    const after = expectAllocated(idOf(body, 'After'));
    expect(valid).not.toBe(after);
    const mapped = new Map(saved.paragraphs.map((entry) => [entry.session.paraId, entry.persisted]));
    expect(mapped.get(atStart.secondParaId)?.paraId).toBe(valid);
    expect(mapped.get(atEnd.secondParaId)?.paraId).toBe(after);
  });

  it('keeps missing, duplicate and malformed source IDs as authored by default', async () => {
    const opened = await open(fixture(), 7);
    opened.insertText({ story: 'body', paraId: 'body:p1', offset: 0 }, 'Still ');
    const saved = await saveYrsDocx(opened);
    const body = savedParagraphs(saved.bytes, DOCUMENT)[DOCUMENT]!;
    expect(idOf(body, 'Still Missing')).toBeNull();
    expect(idOf(body, 'Duplicate')).toBe('1a2b3c4d');
    expect(idOf(body, 'Malformed')).toBe('xyz');
    expect(idOf(body, 'Retained')).toBe('5E6F7A8B');
    const footnotes = savedParagraphs(saved.bytes, FOOTNOTES)[FOOTNOTES]!;
    expect(footnotes.map(([id]) => id)).toEqual([null, null, '3C4D5E6F']);
    expect(saved.paragraphs.map((entry) => entry.session.paraId)).not.toContain('body:p1');
  });

  it('persists every part by occurrence and patches only start tags in an identity-only save', async () => {
    const bytes = fixture();
    const opened = await open(bytes, 7);
    const receipt = opened.persistParagraphIds();
    if (receipt.status !== 'applied') throw new Error('refused');
    expect(receipt.diagnostics).toEqual([]);
    const persisted = new Map(
      receipt.assignments.map((assignment) => [
        assignment.paragraph.kind === 'session'
          ? assignment.paragraph.paraId
          : `${assignment.paragraph.partUri}#${assignment.paragraph.paragraphOrdinal}`,
        assignment,
      ])
    );
    expect([...persisted.keys()]).toEqual([
      'body:p3',
      'body:p1',
      'body:p4',
      '/word/footnotes.xml#0',
      '/word/footnotes.xml#1',
    ]);
    expect(persisted.get('body:p3')).toMatchObject({
      idOrigin: 'repaired',
      previousOoxmlParaId: '1a2b3c4d',
    });
    expect(persisted.get('/word/footnotes.xml#0')?.persisted?.story).toEqual({
      partUri: '/word/footnotes.xml',
      kind: 'footnote',
      itemId: '-1',
    });
    const state = opened.encodeState();
    expect(opened.persistParagraphIds()).toEqual({
      status: 'applied',
      assignments: [],
      diagnostics: [],
    });
    expect(opened.encodeState()).toEqual(state);

    const saved = await saveYrsDocx(opened, { updateModifiedDate: false });
    const id = (owner: string) => expectAllocated(persisted.get(owner)!.ooxmlParaId);
    const original = parts(bytes);
    const written = parts(saved.bytes);
    const decode = (value: Uint8Array | undefined) => new TextDecoder().decode(value);
    const expected = decode(original[DOCUMENT])
      .replace('<w:p><w:r><w:t>Missing', `<w:p w14:paraId="${id('body:p1')}"><w:r><w:t>Missing`)
      .replace('w14:paraId="1a2b3c4d"', `w14:paraId="${id('body:p3')}"`)
      .replace('w14:paraId="xyz"', `w14:paraId="${id('body:p4')}"`);
    expect(decode(written[DOCUMENT])).toBe(expected);
    expect(decode(written[FOOTNOTES])).toBe(
      decode(original[FOOTNOTES])
        .replace('<w:p><w:r><w:separator/>', `<w:p w14:paraId="${id('/word/footnotes.xml#0')}"><w:r><w:separator/>`)
        .replace(
          '<w:p><w:r><w:continuationSeparator/>',
          `<w:p w14:paraId="${id('/word/footnotes.xml#1')}"><w:r><w:continuationSeparator/>`
        )
    );
    expect(Object.keys(written).sort()).toEqual(Object.keys(original).sort());
    for (const name of Object.keys(original)) {
      if (name !== DOCUMENT && name !== FOOTNOTES) expect(written[name]).toEqual(original[name]);
    }
    const all = savedParagraphs(saved.bytes, DOCUMENT, FOOTNOTES, 'word/header1.xml', 'word/comments.xml');
    const ids = Object.values(all).flatMap((paragraphs) => paragraphs.map(([value]) => value));
    expect(ids.every((value) => value !== null)).toBe(true);
    expect(new Set(ids.map((value) => value!.toUpperCase())).size).toBe(ids.length);

    const reopened = await open(saved.bytes, 8);
    const separator = persisted.get('/word/footnotes.xml#0')!.persisted!;
    expect(found(reopened, separator)).toMatchObject({ kind: 'source', paragraphOrdinal: 0 });
    expect(resolvedText(reopened, persisted.get('body:p1')!.persisted!)).toBe('Missing');
    expect(resolvedText(reopened, { kind: 'persisted', story: BODY, paraId: '1A2B3C4D' })).toBe(
      'Valid'
    );
  });

  it('applies persisted IDs through a full save of an edited story', async () => {
    const opened = await open(fixture(), 7);
    const receipt = opened.persistParagraphIds();
    if (receipt.status !== 'applied') throw new Error('refused');
    opened.insertText({ story: 'body', paraId: '0000abcd', offset: 0 }, 'Edited ');
    const saved = await saveYrsDocx(opened);
    const written = savedParagraphs(saved.bytes, DOCUMENT, FOOTNOTES);
    const byOwner = new Map(
      receipt.assignments.map((assignment) => [
        assignment.paragraph.kind === 'session' ? assignment.paragraph.paraId : '',
        assignment.ooxmlParaId,
      ])
    );
    const body = written[DOCUMENT]!;
    expect(idOf(body, 'Missing')).toBe(byOwner.get('body:p1')!);
    expect(idOf(body, 'Duplicate')).toBe(byOwner.get('body:p3')!);
    expect(idOf(body, 'Malformed')).toBe(byOwner.get('body:p4')!);
    expect(idOf(body, 'Edited Lower')).toBe('0000abcd');
    expect(idOf(body, 'Retained')).toBe('5E6F7A8B');
    expect(written[FOOTNOTES]!.every(([id]) => id !== null)).toBe(true);
    expect(text(saved.bytes, DOCUMENT)).not.toContain('"xyz"');
  });

  it('refuses persistence that would guess at an ambiguous comment reference', async () => {
    const opened = await open(
      fixture([
        'word/comments.xml',
        '</w:comments>',
        '<w:comment w:id="2" w:author="Reviewer"><w:p w14:paraId="4D5E6F7A"><w:r><w:t>Twin</w:t></w:r></w:p></w:comment></w:comments>',
      ]),
      7
    );
    const state = opened.encodeState();
    expect(opened.persistParagraphIds()).toEqual({
      status: 'refused',
      refusal: {
        kind: 'ambiguous-comment-reference',
        ooxmlParaId: '4D5E6F7A',
        commentIds: ['1', '2'],
      },
    });
    expect(opened.encodeState()).toEqual(state);
  });

  it('carries identities through content-control values and maps only what is written', async () => {
    const opened = await open(fixture([DOCUMENT, TAIL, `${DROPDOWN}${TAIL}`]), 7);
    const child = opened
      .storyIds()
      .find((story) => opened.paragraphs(story).some((paragraph) => paragraph.paraId === '3A3B3C3D'))!;
    opened.setContentControlValue('42', { kind: 'dropdown', value: 'b' });
    const saved = await saveYrsDocx(opened);
    const body = savedParagraphs(saved.bytes, DOCUMENT)[DOCUMENT]!;
    expect(idOf(body, 'Beta')).toBe('3A3B3C3D');
    expect(body.map(([id]) => id)).not.toContain('3A3B3C3E');
    const sdt = saved.paragraphs.filter((entry) => entry.session.story === child);
    expect(sdt).toEqual([
      {
        session: { ...sdt[0]!.session, paraId: '3A3B3C3D' },
        persisted: { kind: 'persisted', story: BODY, paraId: '3A3B3C3D' },
      },
    ]);
    const written = new Set(body.map(([id]) => id));
    for (const { persisted } of saved.paragraphs) {
      if (persisted.story.partUri === BODY.partUri) expect(written.has(persisted.paraId)).toBe(true);
    }
  });

  it('gives copies new IDs while the original keeps its own', async () => {
    const opened = await open(fixture(), 7);
    opened.applyRawOps('body', [
      {
        op: 'insertEmbed',
        index: 0,
        kind: 'pilcrow',
        payload: { paraId: 'copy', ooxmlParaId: '2b3c4d5e', sourceParaId: '2B3C4D5E' },
      },
    ]);
    const cell: DocxPersistedParagraphAnchor = { kind: 'persisted', story: BODY, paraId: '2B3C4D5E' };
    expect(resolvedText(opened, cell)).toBe('Cell');
    const lower = opened.locateParagraph('body', '0000abcd').start;
    opened.applyRawOps('body', [
      {
        op: 'insertEmbed',
        index: lower,
        kind: 'pilcrow',
        payload: { paraId: '0000abcd', ooxmlParaId: '0000abcd' },
      },
    ]);
    expect(key(opened, 'body', 'Lower')).toBe('0000abcd');
    const split = opened.splitParagraph({ story: 'body', paraId: 'body:p4', offset: 3 });
    const authored = opened
      .paragraphIdentities()
      .paragraphs.find((identity) => identity.session?.paraId === split.secondParaId)!;
    opened.applyRawOps('body', [
      {
        op: 'insertEmbed',
        index: opened.locateParagraph('body', split.secondParaId).start,
        kind: 'pilcrow',
        payload: { paraId: split.secondParaId, ooxmlParaId: authored.ooxmlParaId },
      },
    ]);
    expect(key(opened, 'body', 'formed')).toBe(split.secondParaId);
    const saved = await saveYrsDocx(opened);
    const body = savedParagraphs(saved.bytes, DOCUMENT)[DOCUMENT]!;
    expect(idOf(body, 'Cell')).toBe('2B3C4D5E');
    expect(idOf(body, 'Lower')).toBe('0000abcd');
    expect(idOf(body, 'formed')).toBe(authored.ooxmlParaId);
    const copies = [body[0]![0], body[body.findIndex(([, text]) => text === 'Lower') - 1]![0]];
    for (const copy of copies) expect(expectAllocated(copy)).not.toMatch(/^(2B3C4D5E|0000ABCD)$/);
    const reopened = await open(saved.bytes, 8);
    expect(resolvedText(reopened, cell)).toBe('Cell');
    expect(resolvedText(reopened, { kind: 'persisted', story: BODY, paraId: '0000abcd' })).toBe(
      'Lower'
    );
  });

  it('allocates an ID for an editor-only tail once content is authored into it', async () => {
    const bytes = fixture([
      DOCUMENT,
      TAIL,
      '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId="0A0B0C0D"><w:r><w:t>Last cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    ]);
    const opened = await open(bytes, 7);
    const tail = opened.paragraphs('body').at(-1)!.paraId;
    const sentinel = opened
      .paragraphIdentities()
      .paragraphs.find((identity) => identity.session?.paraId === tail)!;
    expect(sentinel).toMatchObject({ origin: 'synthetic', ooxmlParaId: null, source: null });
    const receipt = opened.persistParagraphIds();
    if (receipt.status !== 'applied') throw new Error('refused');
    expect(
      receipt.assignments.some(
        (assignment) => assignment.paragraph.kind === 'session' && assignment.paragraph.paraId === tail
      )
    ).toBe(false);

    const span = opened.locateParagraph('body', tail);
    opened.insertText({ story: 'body', paraId: tail, offset: span.end - span.start }, 'Typed');
    const promoted = opened
      .paragraphIdentities()
      .paragraphs.find((identity) => identity.session?.paraId === tail)!;
    expect(promoted).toMatchObject({ origin: 'authored', idOrigin: 'authored' });
    const saved = await saveYrsDocx(opened);
    const id = expectAllocated(idOf(savedParagraphs(saved.bytes, DOCUMENT)[DOCUMENT]!, 'Typed'));
    expect(promoted.ooxmlParaId).toBe(id);
    expect(resolvedText(await open(saved.bytes, 8), promoted.persisted!)).toBe('Typed');
  });

  it('resolves a nested anchor after a table inserted before it shifts story positions', async () => {
    const opened = await open(fixture(), 7);
    opened.insertTable({ story: 'body', paraId: '1A2B3C4D', offset: 0 }, 1, 1);
    const saved = await saveYrsDocx(opened);
    const reopened = await open(saved.bytes, 8);
    const cell = found(reopened, { kind: 'persisted', story: BODY, paraId: '2B3C4D5E' });
    expect(cell).toMatchObject({ kind: 'session', story: 'body:t1:r0c0', paraId: '2B3C4D5E' });
    expect(resolvedText(reopened, cell)).toBe('Cell');
  });

  for (const entry of ['openDocx', 'seedFromDocx'] as const) {
  it(`never resolves a stale session anchor in a fresh ${entry} by the same client`, async () => {
    const seed = async (bytes: Uint8Array) => {
      const created = await session(7);
      if (entry === 'openDocx') created.openDocx(bytes, true);
      else created.seedFromDocx(bytes);
      return created;
    };
    const paragraphs = ['A', 'B', 'C'].map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`);
    const body = text(fixture(), DOCUMENT);
    const content = body.slice(body.indexOf('<w:body>') + 8, body.indexOf('<w:sectPr>'));
    const opened = await seed(fixture([DOCUMENT, content, paragraphs.join('')]));
    const stale: DocxParagraphAnchor = {
      kind: 'session',
      sessionId: opened.paragraphIdentities().sessionId,
      story: 'body',
      paraId: 'body:p2',
    };
    expect(resolvedText(opened, stale)).toBe('C');
    const split = opened.splitParagraph({ story: 'body', paraId: 'body:p0', offset: 1 });
    opened.insertText({ story: 'body', paraId: split.secondParaId, offset: 0 }, 'X');
    const saved = await saveYrsDocx(opened);

    const reopened = await seed(saved.bytes);
    expect(key(reopened, 'body', 'B')).toBe('body:p2');
    expect(reopened.resolveParagraphAnchor(stale)).toEqual({
      status: 'unsupported',
      reason: 'foreign-session',
    });
  });
  }

  it('resolves an ID persisted under a rebound prefix after saving and reopening', async () => {
    const opened = await open(
      fixture([DOCUMENT, '<w:p><w:r><w:t>Missing', '<w:p xmlns:w14="urn:other"><w:r><w:t>Missing']),
      7
    );
    const receipt = opened.persistParagraphIds();
    if (receipt.status !== 'applied') throw new Error('refused');
    const missing = receipt.assignments.find(
      (assignment) => assignment.paragraph.kind === 'session' && assignment.paragraph.paraId === 'body:p1'
    )!;
    const saved = await saveYrsDocx(opened);
    expect(idOf(savedParagraphs(saved.bytes, DOCUMENT)[DOCUMENT]!, 'Missing')).toBe(
      missing.ooxmlParaId
    );
    expect(text(saved.bytes, DOCUMENT)).toContain(`w14p0:paraId="${missing.ooxmlParaId}"`);
    expect(resolvedText(await open(saved.bytes, 8), missing.persisted!)).toBe('Missing');
  });

  it('replaces an invalid ID under an alias prefix when a full save persists one', async () => {
    const alias = `<w:p xmlns:x="http://schemas.microsoft.com/office/word/2010/wordml" x:paraId="xyz">`;
    const opened = await open(
      fixture([DOCUMENT, '<w:p><w:r><w:t>Missing', `${alias}<w:r><w:t>Missing`]),
      7
    );
    const receipt = opened.persistParagraphIds();
    if (receipt.status !== 'applied') throw new Error('refused');
    const missing = receipt.assignments.find(
      (assignment) => assignment.paragraph.kind === 'session' && assignment.paragraph.paraId === 'body:p1'
    )!;
    opened.insertText({ story: 'body', paraId: '0000abcd', offset: 0 }, 'Edited ');
    const saved = await saveYrsDocx(opened);
    expect(idOf(savedParagraphs(saved.bytes, DOCUMENT)[DOCUMENT]!, 'Missing')).toBe(
      missing.ooxmlParaId
    );
    expect(text(saved.bytes, DOCUMENT)).not.toContain('"xyz"');
  });

  for (const copyKey of ['copy', '0000abcd']) {
    it(`keeps a deleted source identity from a copy keyed ${copyKey}`, async () => {
      const opened = await open(fixture(), 7);
      const anchor: DocxPersistedParagraphAnchor = { kind: 'persisted', story: BODY, paraId: '0000abcd' };
      const stale: DocxParagraphAnchor = {
        kind: 'session',
        sessionId: opened.paragraphIdentities().sessionId,
        story: 'body',
        paraId: '0000abcd',
      };
      expect(resolvedText(opened, anchor)).toBe('Lower');
      expect(resolvedText(opened, stale)).toBe('Lower');
      opened.mergeParagraphs('body', 'body:p1');
      expect(opened.resolveParagraphAnchor(anchor)).toEqual({ status: 'missing' });
      opened.applyRawOps('body', [
        {
          op: 'insertEmbed',
          index: 0,
          kind: 'pilcrow',
          payload: { paraId: copyKey, ooxmlParaId: '0000abcd' },
        },
      ]);
      const copy = opened
        .paragraphIdentities()
        .paragraphs.find((identity) => identity.session?.story === 'body')!;
      expect(copy.session!.paraId).not.toBe('0000abcd');
      expect(copy.ooxmlParaId).not.toBe('0000abcd');
      expect(opened.resolveParagraphAnchor(anchor)).toEqual({ status: 'missing' });
      expect(opened.resolveParagraphAnchor(stale)).toEqual({ status: 'missing' });
    });
  }

  it('keeps a saved raw identity reserved after its paragraph is deleted', async () => {
    const opened = await open(fixture(), 7);
    const insert = () =>
      opened.applyRawOps('body', [
        {
          op: 'insertEmbed',
          index: 0,
          kind: 'pilcrow',
          payload: { paraId: 'raw', ooxmlParaId: '12345678' },
        },
      ]);
    insert();
    const stale: DocxParagraphAnchor = {
      kind: 'session',
      sessionId: opened.paragraphIdentities().sessionId,
      story: 'body',
      paraId: 'raw',
    };
    const saved = await saveYrsDocx(opened);
    const persisted = saved.paragraphs.find(({ session: anchor }) => anchor.paraId === 'raw')!
      .persisted;
    expect(persisted).toEqual({ kind: 'persisted', story: BODY, paraId: '12345678' });
    expect(found(opened, persisted)).toEqual(found(opened, stale));

    opened.applyRawOps('body', [{ op: 'delete', index: 0, len: 1 }]);
    insert();
    const replacement = opened
      .paragraphIdentities()
      .paragraphs.find((identity) => identity.session?.story === 'body')!;
    expect(replacement.session!.paraId).not.toBe('raw');
    expect(replacement.ooxmlParaId).not.toBe('12345678');
    expect(opened.resolveParagraphAnchor(stale)).toEqual({ status: 'missing' });
    expect(opened.resolveParagraphAnchor(persisted)).toEqual({ status: 'missing' });
    const reopened = await open(saved.bytes, 8);
    expect(found(reopened, persisted).kind).toBe('session');
  });

  it('saves a comment added to the session with its range', async () => {
    const opened = await open(fixture(), 7);
    const { commentId } = opened.addComment(
      [
        {
          story: 'body',
          start: { paraId: '0000abcd', offset: 1 },
          end: { paraId: '0000abcd', offset: 4 },
        },
      ],
      'Ada',
      '2026-09-25T00:00:00Z',
      'Check this'
    );
    expect(commentId).not.toMatch(/^\d+$/);
    const saved = await saveYrsDocx(opened);
    const reopened = await open(saved.bytes, 8);
    const comments = reopened.materializeDocx()!.package.document.comments!;
    const text = (comment: (typeof comments)[number]) =>
      comment.content
        .flatMap((paragraph) => paragraph.content)
        .flatMap((item) => (item.type === 'run' ? item.content : []))
        .map((item) => (item.type === 'text' ? item.text : ''))
        .join('');
    expect(comments.map((comment) => [comment.id, comment.author, text(comment)])).toEqual([
      [1, 'Reviewer', 'Comment'],
      [2, 'Ada', 'Check this'],
    ]);
    expect(comments[1]!.date).toBe('2026-09-25T00:00:00Z');
    const lower = reopened.locateParagraph('body', '0000abcd').start;
    expect(reopened.resolveComment('2')).toEqual([
      { story: 'body', start: lower + 1, end: lower + 4 },
    ]);
    expect(reopened.resolveComment('1')).toEqual(opened.resolveComment('1'));
  });

  it('saves the range of a comment that covers only embedded content', async () => {
    const opened = await open(fixture(), 7);
    const range = {
      story: 'body',
      start: { paraId: 'body:p1', offset: 7 },
      end: { paraId: 'body:p1', offset: 8 },
    };
    opened.addComment([range], 'Ada', '2026-09-25T00:00:00Z', 'Note');
    const saved = await saveYrsDocx(opened);
    const reopened = await open(saved.bytes, 8);
    const added = reopened
      .materializeDocx()!
      .package.document.comments!.find((comment) => comment.author === 'Ada')!;
    const missing = reopened.locateParagraph('body', 'body:p1').start;
    expect(reopened.resolveComment(String(added.id))).toEqual([
      { story: 'body', start: missing + 7, end: missing + 8 },
    ]);
  });

  it('saves the anchor of a comment at a caret', async () => {
    const opened = await open(fixture(), 7);
    const caret = { paraId: '0000abcd', offset: 2 };
    const { commentId } = opened.addComment(
      [{ story: 'body', start: caret, end: caret }],
      'Ada',
      '2026-09-25T00:00:00Z',
      'Here'
    );
    const [anchor] = opened.resolveComment(commentId);
    expect(anchor!.start).toBe(anchor!.end);
    const saved = await saveYrsDocx(opened);
    const lower = text(saved.bytes, DOCUMENT).match(/<w:p [^>]*0000abcd.*?<\/w:p>/)![0];
    const positions = [
      '>Lo<',
      '<w:commentRangeStart w:id="2"/>',
      '<w:commentRangeEnd w:id="2"/>',
      '<w:commentReference w:id="2"/>',
      '>wer<',
    ].map((needle) => lower.indexOf(needle));
    expect(positions.every((position, index) => position > (positions[index - 1] ?? -1))).toBe(
      true
    );
    expect(text(saved.bytes, 'word/comments.xml')).toContain('w:id="2"');
  });

  it('reports only the header views a shared part holds once saved', async () => {
    const header = 'word/header1.xml';
    const bytes = fixture(
      [
        'word/_rels/document.xml.rels',
        '<Relationship Id="rIdFootnotes"',
        '<Relationship Id="rIdHeaderFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFootnotes"',
      ],
      [
        DOCUMENT,
        '<w:headerReference w:type="default" r:id="rIdHeader"/>',
        '<w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="first" r:id="rIdHeaderFirst"/>',
      ],
      [header, '</w:p></w:hdr>', '</w:p><w:p><w:r><w:t>Plain</w:t></w:r></w:p></w:hdr>']
    );
    const expectResolvable = async (saved: Awaited<ReturnType<typeof saveYrsDocx>>) => {
      const held = savedParagraphs(saved.bytes, header)[header]!.map(([id]) => id);
      const reopened = await open(saved.bytes, 9);
      for (const { persisted } of [...saved.paragraphs, ...saved.conflicts]) {
        if (persisted.story.partUri === `/${header}`) expect(held).toContain(persisted.paraId);
        expect(found(reopened, persisted).kind).toBe('session');
      }
      return held;
    };

    const opened = await open(bytes, 7);
    if (opened.persistParagraphIds().status !== 'applied') throw new Error('refused');
    const saved = await saveYrsDocx(opened);
    const held = await expectResolvable(saved);
    expect(held[0]).toBe('6F7A8B9C');
    const plain = expectAllocated(held[1]!);
    const views = saved.paragraphs.filter(({ session: anchor }) => anchor.story.startsWith('hf:'));
    expect(views.map(({ persisted }) => persisted.paraId).sort()).toEqual(
      ['6F7A8B9C', '6F7A8B9C', plain, plain].sort()
    );

    const edited = await open(bytes, 8);
    if (edited.persistParagraphIds().status !== 'applied') throw new Error('refused');
    const halves = ['hf:rIdHeader', 'hf:rIdHeaderFirst'].map((story, offset) => {
      const split = edited.splitParagraph({ story, paraId: key(edited, story, 'Plain'), offset: offset + 1 });
      return split.secondParaId;
    });
    const diverged = await saveYrsDocx(edited);
    const written = await expectResolvable(diverged);
    const identities = edited.paragraphIdentities().paragraphs;
    const writtenHalves = halves.filter((half) => {
      const id = identities.find((identity) => identity.session?.paraId === half)!.ooxmlParaId!;
      const reported = [...diverged.paragraphs, ...diverged.conflicts].some(
        ({ session: anchor }) => anchor.paraId === half
      );
      expect(reported).toBe(written.includes(id));
      return reported;
    });
    expect(writtenHalves).toHaveLength(1);
  });

  it('keeps a deleted paragraph\'s saved ID from a paragraph an offline replica gave it', async () => {
    const opened = await open(fixture(), 7);
    const offline = await session(8);
    offline.loadState(opened.encodeState());
    const split = opened.splitParagraph({ story: 'body', paraId: '0000abcd', offset: 2 });
    const saved = await saveYrsDocx(opened);
    const anchor = saved.paragraphs.find(
      (entry) => entry.session.paraId === split.secondParaId
    )!.persisted;
    opened.deleteRange({
      story: 'body',
      start: { paraId: '0000abcd', offset: 2 },
      end: { paraId: split.secondParaId, offset: 3 },
    });
    expect(opened.resolveParagraphAnchor(anchor)).toEqual({ status: 'missing' });

    offline.applyRawOps('body', [
      {
        op: 'insertEmbed',
        index: 0,
        kind: 'pilcrow',
        payload: { paraId: 'offline', ooxmlParaId: anchor.paraId },
      },
    ]);
    opened.applyUpdate(offline.encodeStateAsUpdate(opened.encodeStateVector()));
    const claimed = opened
      .paragraphIdentities()
      .paragraphs.find((identity) => identity.session?.paraId === 'offline')!;
    expect(claimed.ooxmlParaId).not.toBe(anchor.paraId);
    expect(claimed.idOrigin).toBe('repaired');
    expect(opened.resolveParagraphAnchor(anchor)).toEqual({ status: 'missing' });
    expect(resolvedText(await open(saved.bytes, 9), anchor)).toBe('wer');
  });

  it('reports saved paragraphs whose IDs the session reassigned while saving', async () => {
    const opened = await open(fixture(), 7);
    const pending = saveYrsDocx(opened);
    const receipt = opened.persistParagraphIds();
    const saved = await pending;
    const repaired =
      receipt.status === 'applied'
        ? receipt.assignments.find(
            (assignment) =>
              assignment.paragraph.kind === 'session' && assignment.paragraph.paraId === 'body:p3'
          )
        : undefined;
    expect(repaired?.previousOoxmlParaId).toBe('1a2b3c4d');
    expect(saved.conflicts).toEqual([
      {
        session: { kind: 'session', sessionId: expect.any(String), story: 'body', paraId: 'body:p3' },
        persisted: { kind: 'persisted', story: BODY, paraId: '1a2b3c4d' },
      },
    ]);
    expect(saved.paragraphs.map((entry) => entry.session.paraId)).not.toContain('body:p3');
    expect(saved.paragraphs.map((entry) => entry.persisted.paraId)).toContain('1A2B3C4D');
    expect(idOf(savedParagraphs(saved.bytes, DOCUMENT)[DOCUMENT]!, 'Duplicate')).toBe('1a2b3c4d');
  });

  it('restores retained inline XML of a paragraph that repeats a source ID', async () => {
    const opened = await open(
      fixture([
        DOCUMENT,
        '<w:p w14:paraId="1a2b3c4d"><w:r><w:t>Duplicate',
        '<w:p w14:paraId="1a2b3c4d"><bofx:mark/><w:r><w:t>Duplicate',
      ]),
      7
    );
    opened.insertText({ story: 'body', paraId: '0000abcd', offset: 0 }, 'Edited ');
    const saved = await saveYrsDocx(opened);
    const xml = text(saved.bytes, DOCUMENT);
    const at = xml.indexOf('<w:t>Duplicate');
    expect(xml.slice(xml.lastIndexOf('<w:p', at), at)).toContain('<bofx:mark/>');
    expect(idOf(savedParagraphs(saved.bytes, DOCUMENT)[DOCUMENT]!, 'Edited Lower')).toBe('0000abcd');
  });

  it('resolves session anchors on every replica of one document session', async () => {
    const bytes = fixture();
    const origin = await open(bytes, 7);
    const replica = await session(8);
    replica.loadState(origin.encodeState());
    replica.openDocx(bytes, false);
    const sessionId = origin.paragraphIdentities().sessionId;
    expect(replica.paragraphIdentities().sessionId).toBe(sessionId);

    const split = replica.splitParagraph({ story: 'body', paraId: '0000abcd', offset: 2 });
    const anchor: DocxParagraphAnchor = {
      kind: 'session',
      sessionId,
      story: 'body',
      paraId: split.secondParaId,
    };
    expect(origin.resolveParagraphAnchor(anchor)).toEqual({ status: 'missing' });
    origin.applyUpdate(replica.encodeStateAsUpdate(origin.encodeStateVector()));
    expect(resolvedText(origin, anchor)).toBe('wer');
    expect(origin.resolveParagraphAnchor({ ...anchor, story: 'hf:rIdHeader' })).toEqual({
      status: 'missing',
    });

    const reopened = await open(bytes, 9);
    expect(reopened.resolveParagraphAnchor(anchor)).toEqual({
      status: 'unsupported',
      reason: 'foreign-session',
    });
    expect(key(reopened, 'body', 'Lower')).toBe('0000abcd');
  });

  it('reports ambiguous source IDs and leaves the document unchanged when reading', async () => {
    const opened = await open(fixture(), 7);
    const state = opened.encodeState();
    const identities = opened.paragraphIdentities();
    for (const identity of identities.paragraphs) {
      const anchor = identity.session ?? identity.source!;
      expect(found(opened, anchor)).toEqual(anchor);
    }
    expect(
      opened.resolveParagraphAnchor({ kind: 'persisted', story: BODY, paraId: '1a2b3c4d' })
    ).toEqual({
      status: 'ambiguous',
      candidates: [
        { kind: 'session', sessionId: identities.sessionId, story: 'body', paraId: '1A2B3C4D' },
        { kind: 'session', sessionId: identities.sessionId, story: 'body', paraId: 'body:p3' },
      ],
    });
    const comment = identities.paragraphs.find(
      (identity) => identity.source?.partUri === '/word/comments.xml'
    )!;
    expect(comment.persisted).toEqual({
      kind: 'persisted',
      story: { partUri: '/word/comments.xml', kind: 'comment', itemId: '1' },
      paraId: '4D5E6F7A',
    });
    expect(found(opened, comment.persisted!)).toEqual(comment.source!);
    expect(opened.encodeState()).toEqual(state);

    const detached = await session(10);
    detached.loadStories([{ storyId: 'body', paragraphs: [{ text: 'Plain' }] }]);
    expect(
      detached.resolveParagraphAnchor({ kind: 'persisted', story: BODY, paraId: '1A2B3C4D' })
    ).toEqual({ status: 'unsupported', reason: 'no-source-package' });
  });
});

const PICTURE = '6A6B6C6D';
const AUTHOR = { name: 'Ada', date: '2026-09-25T00:00:00Z' };
const INLINE_IMAGE =
  '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="914400" cy="914400"/><wp:docPr id="7" name="Picture 7"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="7" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
  '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
const RULE = '7A7B7C7D';
const HORIZONTAL_RULE =
  '<w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">' +
  '<v:rect id="rule-1" style="width:0pt;height:1.5pt" o:hr="t" o:hrstd="t" o:hralign="center" fillcolor="#A0A0A0" stroked="f"/></w:pict>';
const BLOCK_CONTROL =
  '<w:sdt><w:sdtPr><w:alias w:val="Choice"/><w:id w:val="42"/></w:sdtPr><w:sdtContent>' +
  '<w:p w14:paraId="3A3B3C3D"><w:r><w:t>Chosen</w:t></w:r></w:p></w:sdtContent></w:sdt>';

/**
 * The shared fixture with an inline image after "Pic", a horizontal rule
 * after "Rule" and a block content control before the tail.
 */
function withImage(): Uint8Array {
  const parts = fixtureParts();
  for (const [name, from, to] of [
    ['[Content_Types].xml', '<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"'],
    [
      'word/_rels/document.xml.rels',
      '</Relationships>',
      '<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
    ],
    [
      DOCUMENT,
      '<w:p w14:paraId="0000abcd">',
      `<w:p w14:paraId="${PICTURE}"><w:r><w:t>Pic</w:t></w:r><w:r>${INLINE_IMAGE}</w:r></w:p>` +
        `<w:p w14:paraId="${RULE}"><w:r><w:t>Rule</w:t></w:r><w:r>${HORIZONTAL_RULE}</w:r></w:p>` +
        '<w:p w14:paraId="0000abcd">',
    ],
    [DOCUMENT, TAIL, `${BLOCK_CONTROL}${TAIL}`],
  ] as const) {
    const xml = parts.get(name)!;
    expect(xml).toContain(from);
    parts.set(name, xml.replace(from, to));
  }
  const pkg: PartsMap = new Map(
    [...parts].map(([name, xml]) => [name, new TextEncoder().encode(xml)])
  );
  pkg.set('word/media/image1.png', new Uint8Array(objects.png));
  return new Uint8Array(rezipPartsToArrayBuffer(pkg));
}

function paragraphXml(bytes: Uint8Array, paraId: string): string {
  return text(bytes, DOCUMENT).match(new RegExp(`<w:p [^>]*${paraId}.*?</w:p>`))![0];
}

const IMAGE_UNIT = {
  story: 'body',
  start: { paraId: PICTURE, offset: 3 },
  end: { paraId: PICTURE, offset: 4 },
};

describe('source bytes are reused only for stories the session left as seeded', () => {
  it('saves a tracked deletion of an image', async () => {
    const opened = await open(withImage(), 7);
    opened.deleteRange(IMAGE_UNIT, AUTHOR);
    const saved = await saveYrsDocx(opened);
    expect(paragraphXml(saved.bytes, PICTURE)).toMatch(/<w:del [^>]*>.*<w:drawing>/);
    const reopened = await open(saved.bytes, 8);
    const picture = reopened
      .materializeDocx()!
      .package.document.content.find(
        (block) => block.type === 'paragraph' && block.paraId === PICTURE
      );
    expect(
      picture?.type === 'paragraph' &&
        picture.content.some(
          (item) =>
            item.type === 'deletion' &&
            item.content.some(
              (run) =>
                run.type === 'run' && run.content.some((content) => content.type === 'drawing')
            )
        )
    ).toBe(true);
  });

  it('saves a tracked insertion of an image', async () => {
    const opened = await open(withImage(), 7);
    const image = opened
      .storySegments('body')
      .find((segment) => segment.kind === 'embed' && segment.embedKind === 'image');
    if (image?.kind !== 'embed') throw new Error('missing seeded image');
    opened.insertImage({ story: 'body', paraId: '0000abcd', offset: 0 }, image.payload, AUTHOR);
    const saved = await saveYrsDocx(opened);
    expect(paragraphXml(saved.bytes, '0000abcd')).toMatch(/<w:ins [^>]*>.*<w:drawing>/);
  });

  it('saves formatting applied to an embed', async () => {
    const opened = await open(withImage(), 7);
    opened.toggleMark(
      { story: 'body', start: { paraId: RULE, offset: 4 }, end: { paraId: RULE, offset: 5 } },
      { type: 'bold' }
    );
    const saved = await saveYrsDocx(opened);
    expect(paragraphXml(saved.bytes, RULE)).toMatch(/<w:b\/>.*<w:pict/);
  });

  it('saves tracked paragraph-mark deletions and insertions', async () => {
    const opened = await open(withImage(), 7);
    opened.deleteRange(
      {
        story: 'body',
        start: { paraId: PICTURE, offset: 4 },
        end: { paraId: '0000abcd', offset: 0 },
      },
      AUTHOR
    );
    opened.splitParagraph({ story: 'body', paraId: '0000abcd', offset: 2 }, AUTHOR);
    const saved = await saveYrsDocx(opened);
    expect(paragraphXml(saved.bytes, PICTURE)).toMatch(/<w:pPr>.*<w:rPr>.*<w:del /);
    expect(paragraphXml(saved.bytes, '0000abcd')).toMatch(/<w:pPr>.*<w:rPr>.*<w:ins /);
  });

  it('saves a comment covering only an image', async () => {
    const opened = await open(withImage(), 7);
    opened.addComment([IMAGE_UNIT], 'Ada', AUTHOR.date, 'Image');
    const saved = await saveYrsDocx(opened);
    expect(paragraphXml(saved.bytes, PICTURE)).toMatch(
      /<w:commentRangeStart w:id="2"\/>.*<w:drawing>.*<w:commentRangeEnd w:id="2"\/>/
    );
  });

  it('saves a table property change', async () => {
    const opened = await open(withImage(), 7);
    opened.setTableWidth({ story: 'body', tableIndex: 0 }, 5000);
    const saved = await saveYrsDocx(opened);
    expect(text(saved.bytes, DOCUMENT)).toContain('<w:tblW w:w="5000" w:type="dxa"/>');
  });

  it('saves a block content control property change', async () => {
    const opened = await open(withImage(), 7);
    const index = opened
      .storySegments('body')
      .reduce<{ at: number; found: number }>(
        (state, segment) => ({
          at: state.at + (segment.kind === 'text' ? segment.text.length : 1),
          found:
            state.found < 0 && segment.kind === 'embed' && segment.embedKind === 'blockSdt'
              ? state.at
              : state.found,
        }),
        { at: 0, found: -1 }
      ).found;
    const [control] = opened
      .storySegments('body')
      .filter((segment) => segment.kind === 'embed' && segment.embedKind === 'blockSdt');
    if (control?.kind !== 'embed') throw new Error('missing block content control');
    const properties = String(control.payload.rawPropertiesXml);
    opened.applyRawOps('body', [
      { op: 'setEmbedAttr', index, key: 'alias', value: 'Renamed' },
      {
        op: 'setEmbedAttr',
        index,
        key: 'rawPropertiesXml',
        value: properties.replace('w:val="Choice"', 'w:val="Renamed"'),
      },
    ]);
    const saved = await saveYrsDocx(opened);
    expect(text(saved.bytes, DOCUMENT)).toContain('<w:alias w:val="Renamed"/>');
  });

  for (const resolution of ['accept', 'reject'] as const) {
    it(`saves an image deletion then ${resolution}ed as resolved`, async () => {
      const opened = await open(withImage(), 7);
      const { revisionId } = opened.deleteRange(IMAGE_UNIT, AUTHOR);
      if (resolution === 'accept') opened.acceptChange({ revisionId: revisionId! });
      else opened.rejectChange({ revisionId: revisionId! });
      const saved = await saveYrsDocx(opened);
      const picture = paragraphXml(saved.bytes, PICTURE);
      expect(picture).not.toContain('<w:del ');
      expect(picture.includes('<w:drawing>')).toBe(resolution === 'reject');
      const reopened = await open(saved.bytes, 8);
      expect(reopened.listRevisions()).toEqual([]);
    });
  }
});
