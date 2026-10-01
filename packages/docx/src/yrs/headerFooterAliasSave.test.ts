import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { headerFooterAliasGroups } from '../docx/headerFooterAliases';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import type { Document, HeaderFooter, Hyperlink, Paragraph, Run } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import {
  captureSessionSave,
  createYrsSession,
  saveYrsDocx,
  writeSessionSave,
  yrsToDocument,
  type YrsSession,
} from './index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WORD = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

function fixture(
  kind: 'header' | 'footer',
  sections = false,
  order = ['rId8', 'rId9']
): ArrayBuffer {
  const reference = (id: string, type: string) => `<w:${kind}Reference w:type="${type}" r:id="${id}"/>`;
  const first = reference('rId8', 'default');
  const second = reference('rId9', sections ? 'default' : 'first');
  const body = sections
    ? `<w:p><w:pPr><w:sectPr>${first}</w:sectPr></w:pPr><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr>${second}</w:sectPr>`
    : `<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr>${first}${second}<w:titlePg/></w:sectPr>`;
  const root = kind === 'header' ? 'hdr' : 'ftr';
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${WORD}.document.main+xml"/><Override PartName="/word/${kind}1.xml" ContentType="${WORD}.${kind}+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="${RELS}"><Relationship Id="office" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="${RELS}">${order.map((id) => `<Relationship Id="${id}" Type="${R}/${kind}" Target="${id === 'rId8' ? '' : './'}${kind}1.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`,
    [`word/${kind}1.xml`]: `<w:${root} xmlns:w="${W}">\n  <w:p><w:r><w:t>Synthetic story</w:t></w:r></w:p>\n</w:${root}>`,
  };
  return rezipPartsToArrayBuffer(new Map(Object.entries(parts).map(([name, xml]) => [name, toBytes(xml)])));
}

function stories(document: Document, kind: 'header' | 'footer'): Map<string, HeaderFooter> {
  return (kind === 'header' ? document.package.headers : document.package.footers)!;
}

function text(story: HeaderFooter): string {
  return (story.content[0] as Paragraph).content.flatMap((item) =>
    item.type === 'run'
      ? item.content.flatMap((content) => content.type === 'text' ? [content.text] : [])
      : []
  ).join('');
}

async function assertSaved(saved: ArrayBuffer, kind: 'header' | 'footer', expected: string): Promise<void> {
  const xml = new TextDecoder().decode(unzipContainer(new Uint8Array(saved))[`word/${kind}1.xml`]);
  const savedText = [...xml.matchAll(/<w:t(?:\s[^>]*)?>(.*?)<\/w:t>/gs)]
    .map((match) => match[1]).join('');
  expect(savedText).toBe(expected);
  const reopened = await parseDocx(saved, { preloadFonts: false });
  expect([...stories(reopened, kind).values()].map(text)).toEqual([expected, expected]);
}

function assertUntouched(source: ArrayBuffer, saved: ArrayBuffer, kind: 'header' | 'footer'): void {
  const before = unzipContainer(new Uint8Array(source));
  const after = unzipContainer(new Uint8Array(saved));
  for (const name of [`word/${kind}1.xml`, 'word/_rels/document.xml.rels']) {
    expect(after[name]).toEqual(before[name]);
  }
}

function insert(session: YrsSession, id: string, value: string): void {
  const story = `hf:${id}`;
  session.insertText({ story, paraId: session.paragraphs(story)[0]!.paraId, offset: 0 }, value);
}

const SAVE_PATHS = ['saveYrsDocx', 'writeSessionSave', 'writeSessionSave full'] as const;
type SavePath = (typeof SAVE_PATHS)[number];
let nextClientId = 98001;

async function open(source: ArrayBuffer): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: nextClientId++ });
  session.openDocx(new Uint8Array(source), true);
  return session;
}

async function save(
  session: YrsSession,
  path: SavePath,
  storyIds?: ReadonlySet<string>
): Promise<ArrayBuffer> {
  const options = { updateModifiedDate: false };
  if (path === 'saveYrsDocx') {
    return Uint8Array.from((await saveYrsDocx(session, options)).bytes).buffer;
  }
  const base = session.materializeDocx()!;
  const capture = captureSessionSave(session);
  const document = yrsToDocument(session, base, { storyIds });
  const result = await writeSessionSave(
    session, document, capture, base.originalBuffer!, options,
    () => path !== 'writeSessionSave full'
  );
  return Uint8Array.from(result.bytes).buffer;
}

async function assertRoundTrip(saved: ArrayBuffer, kind: 'header' | 'footer', expected: string): Promise<void> {
  await assertSaved(saved, kind, expected);
  const session = await open(saved);
  try {
    for (const id of ['rId8', 'rId9']) {
      expect(session.paragraphs(`hf:${id}`)[0]!.text).toBe(expected);
    }
    expect(new Uint8Array(await save(session, 'saveYrsDocx'))).toEqual(new Uint8Array(saved));
  } finally {
    session.destroy();
  }
}

describe('aliased header and footer saves', () => {
  it('repackDocx saves a hyperlink target edit through a shared header', async () => {
    const parts = unzipContainer(new Uint8Array(fixture('header')));
    parts['word/header1.xml'] = toBytes(`<w:hdr xmlns:w="${W}" xmlns:r="${R}"><w:p><w:hyperlink r:id="link"><w:r><w:t>Synthetic link</w:t></w:r></w:hyperlink></w:p></w:hdr>`);
    parts['word/_rels/header1.xml.rels'] = toBytes(`<Relationships xmlns="${RELS}"><Relationship Id="link" Type="${R}/hyperlink" Target="https://example.test/initial" TargetMode="External"/></Relationships>`);
    const source = rezipPartsToArrayBuffer(new Map(Object.entries(parts)));
    const document = await parseDocx(source, { preloadFonts: false });
    const paragraph = document.package.headers!.get('rId8')!.content[0] as Paragraph;
    (paragraph.content[0] as Hyperlink).href = 'https://example.test/edited';
    const saved = await repackDocx(document);
    const reopened = await parseDocx(saved, { preloadFonts: false });
    for (const story of reopened.package.headers!.values()) {
      const paragraph = story.content[0] as Paragraph;
      expect((paragraph.content[0] as Hyperlink).href).toBe('https://example.test/edited');
    }
  });

  it('groups exact resolved paths across headers and footers and excludes inert relationships', async () => {
    const { package: pkg } = await parseDocx(fixture('header'), { preloadFonts: false });
    const part = pkg.headers!.get('rId8')!;
    pkg.footers = new Map([['footerAlias', part]]);
    pkg.relationships.set('footerAlias', {
      id: 'footerAlias', type: `${R}/footer`, target: '/word/nested/../header1.xml',
    });
    expect(headerFooterAliasGroups(pkg)).toEqual([[
      { kind: 'headers', rId: 'rId8' },
      { kind: 'headers', rId: 'rId9' },
      { kind: 'footers', rId: 'footerAlias' },
    ]]);
    const alias = pkg.relationships.get('rId9')!;
    for (const target of ['Header1.xml', '', '../../header1.xml']) {
      alias.target = target;
      expect(headerFooterAliasGroups(pkg)[0]).toHaveLength(2);
    }
    alias.target = 'header1.xml';
    alias.targetMode = 'External';
    expect(headerFooterAliasGroups(pkg)[0]).toHaveLength(2);
    pkg.relationships.delete('footerAlias');
    expect(headerFooterAliasGroups(pkg)).toEqual([]);
    pkg.relationships.set('missing', { id: 'missing', type: `${R}/header`, target: 'header1.xml' });
    expect(headerFooterAliasGroups(pkg)).toEqual([]);
  });

  for (const kind of ['header', 'footer'] as const) {
    for (const sections of [false, true]) {
      it(`parseDocx/repackDocx shares ${kind} aliases across ${sections ? 'sections' : 'page types'}`, async () => {
        for (const id of ['rId8', 'rId9']) {
          const source = fixture(kind, sections);
          const document = await parseDocx(source, { preloadFonts: false });
          const parts = stories(document, kind);
          expect(parts.get('rId8')).toBe(parts.get('rId9'));
          const paragraph = parts.get(id)!.content[0] as Paragraph;
          const item = (paragraph.content[0] as Run).content[0]!;
          if (item.type !== 'text') throw new Error('text');
          item.text = 'Edited synthetic story';
          await assertRoundTrip(
            await repackDocx(document, { updateModifiedDate: false }), kind, 'Edited synthetic story'
          );
          expect(text(parts.get(id === 'rId8' ? 'rId9' : 'rId8')!)).toBe('Edited synthetic story');
        }
      });

      for (const path of SAVE_PATHS) {
        it(`${path} saves either ${kind} alias across ${sections ? 'sections' : 'page types'}`, async () => {
          for (const id of ['rId8', 'rId9']) {
            const source = fixture(kind, sections);
            const session = await open(source);
            try {
              insert(session, id, 'Edited ');
              await assertRoundTrip(await save(session, path), kind, 'Edited Synthetic story');
              await assertSaved(await save(session, path), kind, 'Edited Synthetic story');
            } finally {
              session.destroy();
            }
          }
        });
      }
    }

    for (const path of SAVE_PATHS) {
      it(`${path} saves the last edited ${kind} relationship in both alias orders`, async () => {
        for (const order of [['rId8', 'rId9'], ['rId9', 'rId8']]) {
          for (const edits of [order, [...order].reverse()]) {
            const session = await open(fixture(kind, false, order));
            try {
              for (const id of edits) insert(session, id, `${id} `);
              await assertRoundTrip(
                await save(session, path, new Set([`hf:${order[0]}`])),
                kind, `${order[1]} Synthetic story`
              );
            } finally {
              session.destroy();
            }
          }
        }
      });
    }

    it(`a remote replica saves an edit received through either ${kind} alias`, async () => {
      for (const id of ['rId8', 'rId9']) {
        const source = fixture(kind);
        const session = await open(source);
        const peer = await createYrsSession({ clientId: nextClientId++ });
        try {
          peer.openDocx(new Uint8Array(source), false);
          peer.loadState(session.encodeState());
          insert(session, id, 'Remote ');
          peer.applyUpdate(session.encodeStateAsUpdate(peer.encodeStateVector()));
          for (const path of SAVE_PATHS) {
            await assertRoundTrip(
              await save(peer, path, new Set(['body'])), kind, 'Remote Synthetic story'
            );
          }
        } finally {
          session.destroy();
          peer.destroy();
        }
      }
    });

    it(`keeps the existing ${kind} projection when the edited-story signal is unavailable`, async () => {
      const session = await open(fixture(kind));
      try {
        insert(session, 'rId8', 'Edited ');
        const projected = yrsToDocument(
          { ...session, openRevision: undefined }, session.materializeDocx()!,
          { storyIds: new Set(['hf:rId8']) }
        );
        expect(text(stories(projected, kind).get('rId8')!)).toBe('Edited Synthetic story');
        expect(text(stories(projected, kind).get('rId9')!)).toBe('Synthetic story');
      } finally {
        session.destroy();
      }
    });

    for (const path of ['saveYrsDocx', 'writeSessionSave'] as const) {
      it(`${path} keeps an aliased ${kind} part byte-identical after a body-only edit`, async () => {
        const source = fixture(kind);
        const session = await open(source);
        try {
          session.insertText({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 0 }, 'Body edit ');
          const saved = await save(session, path, new Set(['body']));
          assertUntouched(source, saved, kind);
          expect(new TextDecoder().decode(unzipContainer(new Uint8Array(saved))['word/document.xml'])).toContain('Body edit ');
        } finally {
          session.destroy();
        }
      });
    }

    it(`no-edit saveYrsDocx keeps the shared ${kind} part and its relationships byte-identical`, async () => {
      const source = fixture(kind);
      const session = await open(source);
      try {
        assertUntouched(source, await save(session, 'saveYrsDocx'), kind);
      } finally {
        session.destroy();
      }
    });
  }
});
