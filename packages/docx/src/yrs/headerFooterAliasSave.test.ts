import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
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

function fixture(kind: 'header' | 'footer', sections = false): ArrayBuffer {
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
    'word/_rels/document.xml.rels': `<Relationships xmlns="${RELS}"><Relationship Id="rId8" Type="${R}/${kind}" Target="${kind}1.xml"/><Relationship Id="rId9" Type="${R}/${kind}" Target="./${kind}1.xml"/></Relationships>`,
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

  it('repackDocx writes the last changed entry and leaves shared image bindings untouched', async () => {
    for (const ids of [['rId8', 'rId9'], ['rId9', 'rId8']]) {
      const document = await parseDocx(fixture('header'), { preloadFonts: false });
      const original = document.package.headers!.get('rId8')!;
      const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';
      const story = (value: string, count: number): HeaderFooter => {
        const part = structuredClone(original);
        const run = (part.content[0] as Paragraph).content[0] as Run;
        run.content = [{ type: 'text', text: value }, ...Array.from({ length: count }, () => ({
          type: 'drawing' as const,
          image: { type: 'image' as const, rId: '', src: png, size: { width: 9525, height: 9525 }, wrap: { type: 'inline' as const } },
        }))];
        return part;
      };
      for (const id of ids) {
        document.package.headers!.set(id, story(id === 'rId8' ? 'First entry' : 'Last entry', id === 'rId8' ? 2 : 1));
      }
      const saved = await repackDocx(document);
      const parts = unzipContainer(new Uint8Array(saved));
      expect(Object.keys(parts).filter((path) => path.startsWith('word/media/'))).toHaveLength(1);
      const xml = new TextDecoder().decode(parts['word/header1.xml']);
      expect(xml).toContain('Last entry');
      expect(xml).not.toContain('First entry');
      for (const part of document.package.headers!.values()) {
        const images = (part.content[0] as Paragraph).content
          .flatMap((item) => item.type === 'run' ? item.content : [])
          .filter((item) => item.type === 'drawing');
        expect(images.every((item) => item.type === 'drawing' && item.image.rId === '')).toBe(true);
      }
      const reopened = await parseDocx(saved, { preloadFonts: false });
      for (const part of reopened.package.headers!.values()) {
        const drawing = (part.content[0] as Paragraph).content
          .flatMap((item) => item.type === 'run' ? item.content : [])
          .find((item) => item.type === 'drawing');
        expect(drawing?.type === 'drawing' && drawing.image.src).toBe(png);
      }
    }
  });

  for (const kind of ['header', 'footer'] as const) {
    for (const sections of [false, true]) {
      it(`parseDocx/repackDocx shares ${kind} aliases across ${sections ? 'sections' : 'page types'}`, async () => {
        for (const id of ['rId8', 'rId9']) {
          const source = fixture(kind, sections);
          const document = await parseDocx(source, { preloadFonts: false });
          const parts = stories(document, kind);
          const paragraph = parts.get(id)!.content[0] as Paragraph;
          const item = (paragraph.content[0] as Run).content[0]!;
          if (item.type !== 'text') throw new Error('text');
          item.text = 'Edited synthetic story';
          await assertSaved(await repackDocx(document), kind, 'Edited synthetic story');
          expect(text(parts.get(id === 'rId8' ? 'rId9' : 'rId8')!)).toBe('Edited synthetic story');
          const clean = await parseDocx(source, { preloadFonts: false });
          assertUntouched(source, await repackDocx(clean), kind);
        }
      });

      for (const path of ['saveYrsDocx', 'writeSessionSave', 'writeSessionSave full'] as const) {
        it(`${path} saves either ${kind} alias across ${sections ? 'sections' : 'page types'}`, async () => {
          for (const id of ['rId8', 'rId9']) {
            const source = fixture(kind, sections);
            const session = await createYrsSession({ clientId: 98001 });
            try {
              session.openDocx(new Uint8Array(source), true);
              const save = async (target = session): Promise<ArrayBuffer> => {
                if (path === 'saveYrsDocx') return (await saveYrsDocx(target)).bytes.slice().buffer as ArrayBuffer;
                const base = target.materializeDocx()!;
                return (await writeSessionSave(target, yrsToDocument(target, base), captureSessionSave(target), source, {}, () => path !== 'writeSessionSave full')).bytes.slice().buffer as ArrayBuffer;
              };
              insert(session, id, 'Edited ');
              await assertSaved(await save(), kind, 'Edited Synthetic story');
              await assertSaved(await save(), kind, 'Edited Synthetic story');
              const clean = await createYrsSession({ clientId: 98003 });
              try {
                clean.openDocx(new Uint8Array(source), true);
                assertUntouched(source, await save(clean), kind);
              } finally {
                clean.destroy();
              }
            } finally {
              session.destroy();
            }
          }
        });
      }
    }

    it(`saveYrsDocx deterministically saves the last ${kind} relationship in either edit order`, async () => {
      for (const ids of [['rId8', 'rId9'], ['rId9', 'rId8']]) {
        const source = fixture(kind);
        const session = await createYrsSession({ clientId: 98002 });
        const peer = await createYrsSession({ clientId: 98004 });
        try {
          session.openDocx(new Uint8Array(source), true);
          peer.openDocx(new Uint8Array(source), false);
          peer.loadState(session.encodeState());
          for (const id of ids) {
            insert(session, id!, id === 'rId8' ? 'First ' : 'Last ');
            peer.applyUpdate(session.encodeStateAsUpdate(peer.encodeStateVector()));
          }
          for (const replica of [session, peer]) {
            const saved = (await saveYrsDocx(replica)).bytes.slice().buffer as ArrayBuffer;
            await assertSaved(saved, kind, 'Last Synthetic story');
            expect(new TextDecoder().decode(unzipContainer(new Uint8Array(saved))[`word/${kind}1.xml`])).not.toContain('First');
          }
        } finally {
          session.destroy();
          peer.destroy();
        }
      }
    });
  }
});
