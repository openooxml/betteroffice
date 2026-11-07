import { afterEach, beforeAll, describe, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { engineKind } from '../wasm/enginePng';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { createYrsSession, type YrsSession } from './index';
import fixture from '../../../../crates/docx-edit/tests/fixtures/behind_objects.json';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';

const root = resolve(import.meta.dir, '../../../..');
const deck = new Uint8Array(readFileSync(resolve(root, 'apps/demo/public/betteroffice-demo.docx')));
const native = engineKind === 'native';

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
});

let restore: (() => void) | undefined;
afterEach(() => { restore?.(); restore = undefined; });

function stubBrowser() {
  const calls = { sizes: [] as [number, number][], fills: 0, encoded: 0 };
  class Canvas {
    constructor(public width: number, public height: number) {}
    getContext() {
      if (calls.sizes.at(-1)?.[0] !== this.width || calls.sizes.at(-1)?.[1] !== this.height) calls.sizes.push([this.width, this.height]);
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'fill') return () => { calls.fills++; };
          if (property === 'measureText') return (text: string) => ({ width: text.length });
          if (property === 'createLinearGradient' || property === 'createRadialGradient' || property === 'createPattern') return () => ({ addColorStop: () => undefined });
          return property in target ? target[property as string] : () => undefined;
        },
        set(target, property, value) { target[property as string] = value; return true; },
      });
    }
    convertToBlob(request: { type: string }) { calls.encoded++; return Promise.resolve(new Blob(['png'], { type: request.type })); }
  }
  class Path { moveTo() {} lineTo() {} quadraticCurveTo() {} bezierCurveTo() {} closePath() {} addPath() {} }
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, Path2D: scope.Path2D, Image: scope.Image, createImageBitmap: scope.createImageBitmap };
  Object.assign(scope, {
    OffscreenCanvas: Canvas, Path2D: Path, Image: undefined,
    createImageBitmap: () => Promise.resolve({ width: 4, height: 4, close: () => undefined }),
  });
  restore = () => { Object.assign(scope, saved); };
  return calls;
}

const pngSize = (bytes: Uint8Array): [number, number] => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return [view.getUint32(16), view.getUint32(20)];
};

async function opened(): Promise<YrsSession> {
  const session = await createYrsSession();
  session.openDocx(deck, true);
  return session;
}

describe('YrsSession.renderPng', () => {
  test('draws an EMF picture that reaches the export as its media part', async () => {
    stubBrowser();
    const metafile = new Uint8Array(readFileSync(resolve(root, 'crates/ooxml-metafile/tests/fixtures/clip-bitmap.emf')));
    const image = fixture.image.replace('IMAGE_RANK', '0').replace('IMAGE_BEHIND', '0');
    const parts: PartsMap = new Map([
      ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/x-emf"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
      ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
      ['word/_rels/document.xml.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/></Relationships>')],
      ['word/document.xml', toBytes(`<w:document ${fixture.namespaces}><w:body>${image}${fixture.body}</w:body></w:document>`)],
      ['word/media/image.png', metafile],
    ]);
    const session = await createYrsSession();
    try {
      session.openDocx(new Uint8Array(rezipPartsToArrayBuffer(parts)), true);
      const exported = session.renderPng({ page: 0 });
      if (native) await exported;
      else expect(await exported.then(() => undefined, (error: { code?: string }) => error.code)).toBe('canvas-unavailable');
    } finally { session.destroy(); }
  });

  test('sizes a page by scale and lays out once per version, leaving the session as it was', async () => {
    const calls = stubBrowser();
    const session = await opened();
    try {
      const state = spyOn(session, 'encodeState');
      const version = session.version();
      const one = await session.renderPng({ page: 0 });
      const two = await session.renderPng({ page: 0 }, { scale: 2 });
      const [[w1, h1], [w2, h2]] = native ? [pngSize(one), pngSize(two)] : calls.sizes as [[number, number], [number, number]];
      expect(w1).toBeGreaterThan(100);
      expect(Math.abs(w2! - 2 * w1!)).toBeLessThanOrEqual(1);
      expect(Math.abs(h2! - 2 * h1!)).toBeLessThanOrEqual(1);
      expect(state).toHaveBeenCalledTimes(1);
      expect(session.version()).toBe(version);
    } finally { session.destroy(); }
  });

  test('lays out what the session holds now, not its source', async () => {
    stubBrowser();
    const session = await opened();
    const pages = async () => {
      const refused = await session.renderPng({ page: 9999 }).catch((error: Error) => error.message);
      return Number(/for (\d+) page/.exec(String(refused))?.[1]);
    };
    try {
      const before = await pages();
      expect(before).toBeGreaterThan(0);
      session.insertText({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 0 }, 'A great many more words at the start of this page. '.repeat(75));
      expect(await pages()).toBeGreaterThan(before + 2);
      await session.renderPng({ page: before + 1 });
    } finally { session.destroy(); }
  });

  test('wraps text in a family with no bundled face of its own, as the editor does', async () => {
    stubBrowser();
    const session = await opened();
    try {
      session.insertText({ story: 'body', paraId: session.paragraphs('body')[0]!.paraId, offset: 0 }, 'A great many more words at the start of this page. '.repeat(300));
      const refused = await session.renderPng({ page: 9999 }).catch((error: Error) => error.message);
      expect(Number(/for (\d+) page/.exec(String(refused))?.[1])).toBeGreaterThan(20);
      const png = await session.renderPng({ page: 0 });
      if (native) expect(pngSize(png)[0]).toBeGreaterThan(100);
    } finally { session.destroy(); }
  });

  test('refuses what the native renderer refuses, with the same messages', async () => {
    stubBrowser();
    const session = await opened();
    try {
      await expect(session.renderPng({ page: 9999 })).rejects.toThrow(/^page index 9999 out of range for \d+ page\(s\)$/);
      await expect(session.renderPng({ page: -1 })).rejects.toThrow('page must be a non-negative integer');
      await expect(session.renderPng({ page: 0, extra: 1 } as never)).rejects.toThrow('target must be { page } with a page index');
      await expect(session.renderPng({ page: 0 }, { scale: 0 })).rejects.toThrow('scale must be finite and positive');
      await expect(session.renderPng({ page: 0 }, { scale: 100 })).rejects.toMatchObject({ name: 'PngExportError', code: 'canvas-too-large' });
    } finally { session.destroy(); }
  });

  test('a session destroyed during an export rejects it, and one without source bytes is refused', async () => {
    stubBrowser();
    const session = await opened();
    const exported = session.renderPng({ page: 0 });
    session.destroy();
    await expect(exported).rejects.toThrow('YrsSession is destroyed');
    await expect(session.renderPng({ page: 0 })).rejects.toThrow('YrsSession is destroyed');
    const blank = await createYrsSession();
    try { await expect(blank.renderPng({ page: 0 })).rejects.toThrow('session opened from DOCX bytes'); }
    finally { blank.destroy(); }
  });
});
