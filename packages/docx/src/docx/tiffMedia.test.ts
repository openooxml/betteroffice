import { describe, expect, test } from 'bun:test';
import { MAX_TIFF_BYTES } from '../../../../shared/media';
import type { LayoutBlock } from '../layout/pagination/types';
import { parseDocx } from '.';
import { repackDocx } from './rezip';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { maybeTranscodeTiffMedia } from './tiffMedia';
import { readDocxContainer } from './zipContainer';

const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
const PNG_MAGIC = [137, 80, 78, 71, 13, 10, 26, 10];

function pngBytes(): Uint8Array {
  const binary = atob(PNG_BASE64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

/** Minimal single-strip little-endian 2x1 RGBA TIFF, uncompressed. */
function tinyRgbaTiff(): Uint8Array {
  const pixels = [255, 0, 0, 255, 0, 128, 255, 64];
  const entries: Array<[number, number, number, number]> = [
    [256, 3, 1, 2],
    [257, 3, 1, 1],
    [258, 3, 4, 122],
    [259, 3, 1, 1],
    [262, 3, 1, 2],
    [273, 4, 1, 130],
    [277, 3, 1, 4],
    [278, 4, 1, 1],
    [279, 4, 1, 8],
  ];
  const tiff = new Uint8Array(138);
  const view = new DataView(tiff.buffer);
  view.setUint16(0, 0x4949, true);
  view.setUint16(2, 42, true);
  view.setUint32(4, 8, true);
  view.setUint16(8, entries.length, true);
  entries.forEach(([tag, type, count, value], index) => {
    const offset = 10 + index * 12;
    view.setUint16(offset, tag, true);
    view.setUint16(offset + 2, type, true);
    view.setUint32(offset + 4, count, true);
    view.setUint32(offset + 8, value, true);
  });
  for (let index = 0; index < 4; index += 1) view.setUint16(122 + index * 2, 8, true);
  tiff.set(pixels, 130);
  return tiff;
}

function failingTranscode(): (bytes: Uint8Array) => Uint8Array {
  return () => {
    throw new Error('transcode must not run');
  };
}

describe('maybeTranscodeTiffMedia', () => {
  test('passes non-TIFF media through with mime and dataUrl intact', () => {
    const bytes = pngBytes();
    const dataUrl = `data:image/png;base64,${PNG_BASE64}`;
    expect(maybeTranscodeTiffMedia(bytes, 'image/png', dataUrl, failingTranscode())).toEqual({
      bytes,
      mimeType: 'image/png',
      dataUrl,
    });
  });

  test('rejects TIFF past the transfer budget before transcoding', () => {
    const bytes = new Uint8Array(MAX_TIFF_BYTES + 1);
    bytes.set([0x4d, 0x4d, 0, 0x2a]);
    expect(() =>
      maybeTranscodeTiffMedia(bytes, 'image/tiff', 'data:image/tiff;base64,', failingTranscode())
    ).toThrow('TIFF image exceeds the browser transfer budget');
  });

  test('transcodes TIFF media to PNG bytes', () => {
    const tiff = tinyRgbaTiff();
    const seen: Uint8Array[] = [];
    const result = maybeTranscodeTiffMedia(
      tiff,
      'image/tiff',
      'data:image/tiff;base64,',
      (bytes) => {
        seen.push(bytes.slice());
        return pngBytes();
      }
    );
    expect(seen).toHaveLength(1);
    expect(seen[0]).toEqual(tiff);
    expect(Array.from(result.bytes.subarray(0, 8))).toEqual(PNG_MAGIC);
    expect(result.mimeType).toBe('image/png');
    expect(result.dataUrl).toBe(`data:image/png;base64,${PNG_BASE64}`);
  });
});

let facade: typeof import('./rustParseFacade') | undefined;
try {
  facade = await import('./rustParseFacade');
} catch {
  facade = undefined;
}
const describeIfWasm = facade ? describe : describe.skip;

function mediaEntry(
  alias: string,
  path: string,
  mimeType: string,
  bytes: Uint8Array
): [string, Record<string, string>] {
  const base64 = Buffer.from(bytes).toString('base64');
  return [alias, { path, mimeType, base64, dataUrl: `data:${mimeType};base64,${base64}` }];
}

function envelopeWithMedia(mediaEntries: unknown[]): unknown {
  return {
    wireVersion: 1,
    document: {
      package: {
        document: { content: [] },
        theme: {},
        numbering: {},
        settings: {},
        fontTable: {},
        relationshipEntries: [],
        mediaEntries,
        chartEntries: [],
      },
    },
    embeddedFontParts: [],
  };
}

function pngMagicOf(dataUrl: string): number[] {
  const base64 = dataUrl.replace('data:image/png;base64,', '');
  const binary = atob(base64);
  return Array.from(binary.slice(0, 8)).map((char) => char.charCodeAt(0));
}

describeIfWasm('decodeS9EnvelopeValue TIFF media', () => {
  test('transcodes shared TIFF sources to PNG and leaves other media alone', () => {
    if (!facade) throw new Error('parse wasm is unavailable');
    const tiff = tinyRgbaTiff();
    const png = pngBytes();
    const result = facade.decodeS9EnvelopeValue(
      envelopeWithMedia([
        mediaEntry('word/media/image1.tiff', 'word/media/image1.tiff', 'image/tiff', tiff),
        mediaEntry('media/image1.tiff', 'word/media/image1.tiff', 'image/tiff', tiff),
        mediaEntry('word/media/image2.png', 'word/media/image2.png', 'image/png', png),
      ]),
      new ArrayBuffer(0)
    );
    const media = result.document.package.media;
    expect(media).toBeDefined();
    if (!media) throw new Error('media must be decoded');
    const first = media.get('word/media/image1.tiff');
    const second = media.get('media/image1.tiff');
    expect(first).toBeDefined();
    expect(first).toBe(second);
    expect(first?.mimeType).toBe('image/png');
    expect((first?.dataUrl ?? '').startsWith('data:image/png;base64,')).toBe(true);
    expect(first ? pngMagicOf(first.dataUrl ?? '') : []).toEqual(PNG_MAGIC);
    expect(first ? Array.from(new Uint8Array(first.data).subarray(0, 8)) : []).toEqual(PNG_MAGIC);
    const passthrough = media.get('word/media/image2.png');
    expect(passthrough?.mimeType).toBe('image/png');
    expect(passthrough ? Array.from(new Uint8Array(passthrough.data)) : []).toEqual(
      Array.from(png)
    );
  });
});

/** Minimal single-strip little-endian 2x1 RGB TIFF, uncompressed. */
function tinyRgbTiff(): Uint8Array {
  const entries: Array<[number, number, number, number]> = [
    [256, 3, 1, 2],
    [257, 3, 1, 1],
    [258, 3, 3, 122],
    [259, 3, 1, 1],
    [262, 3, 1, 2],
    [273, 4, 1, 128],
    [277, 3, 1, 3],
    [278, 4, 1, 1],
    [279, 4, 1, 6],
  ];
  const tiff = new Uint8Array(134);
  const view = new DataView(tiff.buffer);
  view.setUint16(0, 0x4949, true);
  view.setUint16(2, 42, true);
  view.setUint32(4, 8, true);
  view.setUint16(8, entries.length, true);
  entries.forEach(([tag, type, count, value], index) => {
    const offset = 10 + index * 12;
    view.setUint16(offset, tag, true);
    view.setUint16(offset + 2, type, true);
    view.setUint32(offset + 4, count, true);
    view.setUint32(offset + 8, value, true);
  });
  for (let index = 0; index < 3; index += 1) view.setUint16(122 + index * 2, 8, true);
  tiff.set([0xff, 0x00, 0x00, 0x00, 0x80, 0xff], 128);
  return tiff;
}

const TIFF_PART = 'word/media/image1.tif';

const TIFF_CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Default Extension="tif" ContentType="image/tiff"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>';

const TIFF_PACKAGE_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

const TIFF_DOCUMENT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.tif"/>' +
  '</Relationships>';

const TIFF_DOCUMENT =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<w:body><w:p><w:r><w:drawing><wp:inline>' +
  '<wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>';

function findImages(value: unknown, found: Record<string, unknown>[] = []): Record<string, unknown>[] {
  if (Array.isArray(value)) {
    for (const item of value) findImages(item, found);
  } else if (value && typeof value === 'object') {
    const record = value as Record<string, unknown>;
    if (record.type === 'image') found.push(record);
    for (const child of Object.values(record)) findImages(child, found);
  }
  return found;
}

describeIfWasm('TIFF media end to end', () => {
  test('renders a TIFF picture as PNG while the saved package keeps the TIFF bytes', async () => {
    const tiff = tinyRgbTiff();
    const parts = new Map<string, Uint8Array>();
    parts.set('[Content_Types].xml', toBytes(TIFF_CONTENT_TYPES));
    parts.set('_rels/.rels', toBytes(TIFF_PACKAGE_RELS));
    parts.set('word/_rels/document.xml.rels', toBytes(TIFF_DOCUMENT_RELS));
    parts.set(TIFF_PART, tiff);
    parts.set('word/document.xml', toBytes(TIFF_DOCUMENT));

    const parsed = await parseDocx(rezipPartsToArrayBuffer(parts), { preloadFonts: false });
    const images = findImages(parsed.package.document.content);
    expect(images).toHaveLength(1);
    expect(images[0]?.rId).toBe('rId5');
    expect(images[0]?.mimeType).toBe('image/png');
    expect(String(images[0]?.src).startsWith('data:image/png;base64,')).toBe(true);
    expect(pngMagicOf(String(images[0]?.src))).toEqual(PNG_MAGIC);

    const saved = readDocxContainer(await repackDocx(parsed));
    expect(Array.from(saved.file(TIFF_PART) ?? [])).toEqual(Array.from(tiff));
    expect(saved.paths().some((path) => path.endsWith('.png'))).toBe(false);
  });

  // The renderer paints the yrs-seeded src, which docx-edit parses itself.
  test('seeds the yrs image src the renderer paints as PNG', async () => {
    const { createYrsSession } = await import('../yrs');
    const parts = new Map<string, Uint8Array>();
    parts.set('[Content_Types].xml', toBytes(TIFF_CONTENT_TYPES));
    parts.set('_rels/.rels', toBytes(TIFF_PACKAGE_RELS));
    parts.set('word/_rels/document.xml.rels', toBytes(TIFF_DOCUMENT_RELS));
    parts.set(TIFF_PART, tinyRgbTiff());
    parts.set('word/document.xml', toBytes(TIFF_DOCUMENT));
    const bytes = new Uint8Array(rezipPartsToArrayBuffer(parts));

    const session = await createYrsSession({ clientId: 75112 });
    try {
      session.seedFromDocx(bytes);
      const blocks = session.yrsBlocksForStory('body', {}) as LayoutBlock[];
      const runs = blocks.flatMap((block) => (block.kind === 'paragraph' ? block.runs : []));
      const image = runs.find((run) => run.kind === 'image');
      expect(image?.kind).toBe('image');
      if (image?.kind !== 'image') throw new Error('missing image');
      expect(image.src.startsWith('data:image/png;base64,')).toBe(true);
      expect(pngMagicOf(image.src)).toEqual(PNG_MAGIC);
    } finally {
      session.destroy();
    }
  });
});
