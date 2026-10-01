import { afterEach, beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { computeAnchorPositionsFromYrs } from '../layout/render/displayListAnchors';
import type { DisplayListQueries } from '../layout/render/displayListQueries';
import { createYrsSidebarProjection } from '../layout/render/yrsSidebarProjection';
import type { BlockContent, Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import {
  createYrsInputPositionMap,
  createYrsPositionProjection,
  createYrsSession,
  displayPositionToYrsLoc,
  documentToYrs,
  projectYrsDisplayPosition,
  saveYrsDocx,
  yrsToDocument,
  type YrsSession,
} from '../yrs';
import { headerFooterStory, sessionInternals } from '../yrs/sessionInternals';
import { parseDocx } from './index';
import { collectParts, rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { decodeS9EnvelopeValue } from './rustParseFacade';
import { unzipContainer } from './wasm';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WORD = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;
const sessions: YrsSession[] = [];

beforeAll(() =>
  preloadEditWasm(new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
  )))
);
afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

async function replica(): Promise<YrsSession> {
  const session = await createYrsSession();
  sessions.push(session);
  return session;
}

function fixture(
  target: string | null = 'header1.xml',
  canonicalDisplayed = true,
  footer = false
): Uint8Array {
  const kind = footer ? 'footer' : 'header';
  const root = footer ? 'ftr' : 'hdr';
  const path = `${kind}1.xml`;
  const section = (rId: string) => `<w:sectPr><w:${kind}Reference w:type="default" r:id="${rId}"/></w:sectPr>`;
  const first = canonicalDisplayed || target === null ? 'rId7' : 'rId9';
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${WORD}.document.main+xml"/><Override PartName="/word/${path}" ContentType="${WORD}.${kind}+xml"/></Types>`)],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="${R}/${kind}" Target="${path}"/>${target === null ? '' : `<Relationship Id="rId9" Type="${R}/${kind}" Target="${target}"/>`}</Relationships>`)],
    ['word/document.xml', toBytes(`<w:document ${NS}><w:body><w:p w14:paraId="00000001"><w:pPr>${section(first)}</w:pPr><w:r><w:t>First</w:t></w:r></w:p><w:p w14:paraId="00000002"><w:r><w:t>Second</w:t></w:r></w:p>${section(target === null ? first : 'rId9')}</w:body></w:document>`)],
    [`word/_rels/${path}.rels`, toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdLink" Type="${R}/hyperlink" Target="https://example.test/band" TargetMode="External"/></Relationships>`)],
    [`word/${path}`, toBytes(`<w:${root} ${NS}><w:p w14:paraId="0000E001"><w:r><w:t>Shared</w:t></w:r></w:p></w:${root}>`)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function aliasedRoom(session: YrsSession, bytes = fixture(), footer = false): Document {
  const { document } = session.openDocx(bytes, true);
  const root = 'hf:rId9';
  for (const story of session.storyIds()) {
    if (story === root || story.startsWith(`${root}:`)) session.deleteStory(story);
  }
  sessionInternals(session).setHeaderFooterAliases(JSON.stringify([{
    isHeader: !footer,
    partPath: `word/${footer ? 'footer' : 'header'}1.xml`,
    relationshipIds: ['rId7', 'rId9'],
  }]));
  return document;
}

function expectBandParts(actual: Uint8Array, source: Uint8Array): void {
  const saved = unzipContainer(actual);
  const original = unzipContainer(source);
  const bandOrRels = (name: string) =>
    /(?:^|\/)(?:header|footer)[^/]*\.xml$/.test(name) || name.endsWith('.rels');
  expect(Object.keys(saved).filter(bandOrRels).sort())
    .toEqual(Object.keys(original).filter(bandOrRels).sort());
  for (const name of Object.keys(original).filter(bandOrRels)) {
    expect(saved[name]).toEqual(original[name]);
  }
}

function contentText(blocks: readonly BlockContent[]): string {
  return blocks.flatMap((block) => block.type === 'paragraph' ? block.content : [])
    .flatMap((inline) => inline.type === 'run' ? inline.content : [])
    .map((inline) => inline.type === 'text' ? inline.text : '').join('');
}

function headerText(document: Document, rId: string): string {
  return contentText(document.package.headers!.get(rId)!.content);
}

function xml(bytes: Uint8Array, path: string): string {
  return new TextDecoder().decode(unzipContainer(bytes)[path]);
}

function expectSameParts(actual: Uint8Array, expected: Uint8Array): void {
  const a = unzipContainer(actual);
  const b = unzipContainer(expected);
  expect(Object.keys(a).sort()).toEqual(Object.keys(b).sort());
  for (const name of Object.keys(b)) expect(a[name]).toEqual(b[name]);
}

describe('header/footer aliases', () => {
  for (const footer of [false, true]) {
    const kind = footer ? 'footer' : 'header';
    for (const target of [`${kind}1.xml`, `./${kind}1.xml`]) {
      for (const native of [true, false]) {
        test(`default ${native ? 'native open' : 'TS seed'} preserves ${target}`, async () => {
          const bytes = fixture(target, true, footer);
          const session = await replica();
          session.openDocx(bytes, native);
          if (!native) {
            const document = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
            documentToYrs(session, document);
          }
          expect(session.storyIds().filter((id) => id.startsWith('hf:')).sort())
            .toEqual(['hf:rId7', 'hf:rId9']);
          expect(headerFooterStory(session, 'rId9')).toBe('hf:rId9');
          expect([...sessionInternals(session).headerFooterAliases()]).toEqual([]);
          expect(new TextDecoder().decode(session.encodeState())).not.toContain('hfAliases');
          const unedited = await saveYrsDocx(session, { updateModifiedDate: false });
          expectBandParts(unedited.bytes, bytes);
          expect(unedited.paragraphs.filter(({ session: anchor }) => anchor.story.startsWith('hf:')))
            .toHaveLength(2);
          const paragraph = session.paragraphs('body')[0]!;
          session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Edited ');
          const edited = await saveYrsDocx(session, { updateModifiedDate: false });
          expectBandParts(edited.bytes, bytes);
          expect(edited.paragraphs.filter(({ session: anchor }) => anchor.story.startsWith('hf:')))
            .toHaveLength(2);
        });
      }
    }
  }

  for (const target of ['header1.xml', './header1.xml']) {
    test(`routes, edits and saves a hand-built alias room for ${target}`, async () => {
      const bytes = fixture(target);
      const session = await replica();
      const document = aliasedRoom(session, bytes);
      expect(session.storyIds().filter((id) => id.startsWith('hf:'))).toEqual(['hf:rId7']);
      expect(headerFooterStory(session, 'rId9')).toBe('hf:rId7');
      const aliases = sessionInternals(session).headerFooterAliases();
      expect([...aliases]).toEqual([['rId9', 'rId7']]);
      expect(sessionInternals(session).headerFooterAliases()).toBe(aliases);
      const paragraph = session.paragraphs('hf:rId7')[0]!;
      session.insertText({ story: 'hf:rId7', paraId: paragraph.paraId, offset: 6 }, ' edited');
      expect(sessionInternals(session).headerFooterAliases()).not.toBe(aliases);
      const projected = yrsToDocument(session, document, { storyIds: new Set(['hf:rId7']) });
      expect(headerText(projected, 'rId7')).toBe('Shared edited');
      expect(headerText(projected, 'rId9')).toBe('Shared edited');
      expect(projected.package.headers!.get('rId7')!.content)
        .toBe(projected.package.headers!.get('rId9')!.content);
      expect(collectParts(projected).filter((part) => part.relsPath.includes('header1'))).toHaveLength(1);

      const saved = await saveYrsDocx(session, { updateModifiedDate: false });
      const receipts = saved.paragraphs.filter(({ session: anchor }) => anchor.story.startsWith('hf:'));
      expect(receipts).toHaveLength(1);
      expect(receipts[0]!.session.story).toBe('hf:rId7');
      expect(session.paragraphIdentities().paragraphs.filter(({ session: anchor }) =>
        anchor?.story.startsWith('hf:')
      )).toHaveLength(1);
      const relationships = xml(saved.bytes, 'word/_rels/document.xml.rels');
      expect(relationships).toContain(`Id="rId7" Type="${R}/header" Target="header1.xml"`);
      expect(relationships).toContain(`Id="rId9" Type="${R}/header" Target="${target}"`);
      const body = xml(saved.bytes, 'word/document.xml');
      expect(body.match(/<w:headerReference[^>]*r:id="rId7"/g)).toHaveLength(1);
      expect(body.match(/<w:headerReference[^>]*r:id="rId9"/g)).toHaveLength(1);
      expect(xml(saved.bytes, 'word/header1.xml').match(/edited/g)).toHaveLength(1);
      const reopened = await replica();
      const host = reopened.openDocx(saved.bytes, true);
      expect(reopened.storyIds().filter((id) => id.startsWith('hf:')).sort())
        .toEqual(['hf:rId7', 'hf:rId9']);
      expect([...sessionInternals(reopened).headerFooterAliases()]).toEqual([]);
      expect(reopened.resolveParagraphAnchor(receipts[0]!.persisted)).toMatchObject({
        status: 'found', anchor: { kind: 'session', story: 'hf:rId7' },
      });
      expect(headerText(yrsToDocument(reopened, host.document), 'rId7')).toBe('Shared edited');
      expect(headerText(yrsToDocument(reopened, host.document), 'rId9')).toBe('Shared edited');
      const second = await saveYrsDocx(reopened, { updateModifiedDate: false });
      expectSameParts(second.bytes, saved.bytes);
    });
  }

  test('projects a section-two header hit to the canonical story and back', async () => {
    const session = await replica();
    aliasedRoom(session);
    const projectionFor = (story: string) => createYrsPositionProjection(session, story);
    const target = projectYrsDisplayPosition(
      { position: 4, pageIndex: 1, region: 'header', rId: 'rId9' }, projectionFor, session
    )!;
    expect(target.story).toBe('hf:rId7');
    const map = createYrsInputPositionMap(target.story, session.paragraphSpans(target.story));
    const loc = displayPositionToYrsLoc(map, target.displayPosition)!;
    expect(loc.story).toBe('hf:rId7');
    expect(projectionFor('hf:rId7')!.positionForLoc(loc)).toBe(4);
  });

  test('anchors find displayed alias bands when the canonical rId is unused', async () => {
    const session = await replica();
    aliasedRoom(session, fixture('header1.xml', false));
    const paragraph = session.paragraphs('hf:rId7')[0]!;
    const { commentId } = session.addComment([{
      story: 'hf:rId7',
      start: { paraId: paragraph.paraId, offset: 0 },
      end: { paraId: paragraph.paraId, offset: 3 },
    }], 'Author', '2026-10-01T00:00:00Z', 'Note');
    const projection = createYrsSidebarProjection(session);
    const point = projection.locToDisplayPoint({ story: 'hf:rId7', paraId: paragraph.paraId, offset: 0 });
    expect(point?.hfRid).toBe('rId7');
    const queries = {
      pageCount: () => 1,
      pageSize: () => ({ width: 800, height: 1000 }),
      hfAnchorRects: (_region: string, rId: string) => rId === 'rId9'
        ? [{ pageIndex: 0, x: 20, y: 30, width: 50, height: 20 }] : [],
    } as unknown as DisplayListQueries;
    expect(computeAnchorPositionsFromYrs(
      session, [commentId], [], projection, queries, new Map([['rId7', 'header'], ['rId9', 'header']])
    ).get(`comment-${commentId}`)).toBe(54);
  });

  test('invalidates aliases after legacy stories arrive through state and updates', async () => {
    const session = await replica();
    const document = aliasedRoom(session);
    expect(headerFooterStory(session, 'rId9')).toBe('hf:rId7');
    const joined = await replica();
    joined.loadState(session.encodeState());
    expect(headerFooterStory(joined, 'rId9')).toBe('hf:rId7');
    session.createStory('hf:rId9', 'Legacy');
    expect([...sessionInternals(session).headerFooterAliases()]).toEqual([]);
    expect(headerFooterStory(session, 'rId9')).toBe('hf:rId9');
    const projected = yrsToDocument(session, document);
    expect(headerText(projected, 'rId7')).toBe('Shared');
    expect(headerText(projected, 'rId9')).toBe('Legacy');
    joined.applyUpdate(session.encodeState());
    expect(headerFooterStory(joined, 'rId9')).toBe('hf:rId9');
    joined.loadState(session.encodeState());
    expect([...sessionInternals(joined).headerFooterAliases()]).toEqual([]);
    expect(joined.storyIds().filter((id) => id.startsWith('hf:')).sort())
      .toEqual(['hf:rId7', 'hf:rId9']);
  });

  test('refreshes the alias cache across history, opening and native seeding', async () => {
    const session = await replica();
    aliasedRoom(session);
    const internals = sessionInternals(session);
    let aliases = internals.headerFooterAliases();
    const paragraph = session.paragraphs('hf:rId7')[0]!;
    session.insertText({ story: 'hf:rId7', paraId: paragraph.paraId, offset: 0 }, 'Edit ');
    expect(internals.headerFooterAliases()).not.toBe(aliases);
    aliases = internals.headerFooterAliases();
    expect(session.undo()).toBe(true);
    expect(internals.headerFooterAliases()).not.toBe(aliases);
    expect(headerFooterStory(session, 'rId9')).toBe('hf:rId7');
    aliases = internals.headerFooterAliases();
    expect(session.redo()).toBe(true);
    expect(internals.headerFooterAliases()).not.toBe(aliases);
    aliases = internals.headerFooterAliases();
    session.beginOpening('cache');
    expect(internals.headerFooterAliases()).not.toBe(aliases);
    aliases = internals.headerFooterAliases();
    session.seedFromDocx(fixture());
    expect(internals.headerFooterAliases()).not.toBe(aliases);
    expect(headerFooterStory(session, 'rId9')).toBe('hf:rId9');
  });

  test('declared groups stay inactive while each relationship owns a story', async () => {
    const session = await replica();
    const { document } = session.openDocx(fixture(), true);
    sessionInternals(session).setHeaderFooterAliases(JSON.stringify([{
      isHeader: true, partPath: 'word/header1.xml', relationshipIds: ['rId7', 'rId9'],
    }]));
    expect([...sessionInternals(session).headerFooterAliases()]).toEqual([]);
    expect(headerFooterStory(session, 'rId9')).toBe('hf:rId9');
    const alias = session.paragraphs('hf:rId9')[0]!;
    session.insertText({ story: 'hf:rId9', paraId: alias.paraId, offset: 6 }, ' legacy');
    const projected = yrsToDocument(session, document);
    expect(headerText(projected, 'rId7')).toBe('Shared');
    expect(headerText(projected, 'rId9')).toBe('Shared legacy');
    expect(projected.package.headers!.get('rId7')!.content)
      .not.toBe(projected.package.headers!.get('rId9')!.content);
  });

  test('shares footer content too', async () => {
    const session = await replica();
    const document = aliasedRoom(session, fixture('./footer1.xml', true, true), true);
    expect(session.storyIds().filter((id) => id.startsWith('hf:'))).toEqual(['hf:rId7']);
    expect(headerFooterStory(session, 'rId9')).toBe('hf:rId7');
    const paragraph = session.paragraphs('hf:rId7')[0]!;
    session.insertText({ story: 'hf:rId7', paraId: paragraph.paraId, offset: 6 }, '!');
    const projected = yrsToDocument(session, document, { storyIds: new Set(['hf:rId7']) });
    expect(contentText(projected.package.footers!.get('rId9')!.content)).toBe('Shared!');
    const saved = await saveYrsDocx(session, { updateModifiedDate: false });
    expect(xml(saved.bytes, 'word/footer1.xml')).toContain('!');
    expect(saved.paragraphs.filter(({ session: anchor }) => anchor.story.startsWith('hf:'))).toHaveLength(1);
  });

  test('singleton keeps stories, metadata and every unedited source part', async () => {
    const bytes = fixture(null);
    const session = await replica();
    session.openDocx(bytes, true);
    expect(session.storyIds()).toEqual(['body', 'hf:rId7']);
    expect(headerFooterStory(session, 'rId7')).toBe('hf:rId7');
    expect([...sessionInternals(session).headerFooterAliases()]).toEqual([]);
    expect(new TextDecoder().decode(session.encodeState())).not.toContain('hfAliases');
    const saved = await saveYrsDocx(session, { updateModifiedDate: false });
    expectSameParts(saved.bytes, bytes);
    session.openDocx(fixture(), true);
    expect(headerFooterStory(session, 'rId9')).toBe('hf:rId9');
  });
});

function wire(groups: unknown): unknown {
  return {
    wireVersion: 1,
    document: { package: {
      document: { content: [] }, theme: {}, numbering: {}, settings: {}, fontTable: {},
      relationshipEntries: [], mediaEntries: [], chartEntries: [], headerFooterAliases: groups,
    } },
    embeddedFontParts: [],
  };
}

test('validates alias wire shape without extending the public package', () => {
  const group = { isHeader: true, partPath: 'word/header1.xml', relationshipIds: ['rId7', 'rId9'] };
  const { document } = decodeS9EnvelopeValue(wire([group]), new ArrayBuffer(0));
  expect('headerFooterAliases' in document.package).toBe(false);
  for (const groups of [
    {}, [null], [{ ...group, isHeader: 'true' }], [{ ...group, partPath: 7 }],
    [{ ...group, relationshipIds: ['rId7'] }], [{ ...group, relationshipIds: ['rId7', ''] }],
    [{ ...group, relationshipIds: ['rId7', 9] }], [{ ...group, extra: true }],
  ]) {
    expect(() => decodeS9EnvelopeValue(wire(groups), new ArrayBuffer(0))).toThrow(TypeError);
  }
});
