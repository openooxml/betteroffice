import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { unzipContainer } from '../wasm/opc';
import { createYrsSession, decodeDocxHostJson, saveYrsDocx, yrsToDocument } from './index';
import { createResidentEngineSession } from './residentEngineSession';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;
const PICTURE = `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

function fixture(image: Uint8Array = PNG, extension = 'png', mimeType = 'image/png'): Uint8Array {
  const contentType = extension === 'png'
    ? ''
    : `<Default Extension="${extension}" ContentType="${mimeType}"/>`;
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${contentType}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`)],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${R}/image" Target="media/picture.${extension}"/><Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/></Relationships>`)],
    ['word/_rels/comments.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${R}/image" Target="media/picture.${extension}"/></Relationships>`)],
    ['word/comments.xml', toBytes(`<w:comments ${NS}><w:comment w:id="0" w:author="A" w:date="2024-01-01T00:00:00Z"><w:p>${PICTURE}</w:p></w:comment></w:comments>`)],
    ['word/media/unused.png', new Uint8Array([9, 9, 9])],
    [`word/media/picture.${extension}`, image],
    ['word/document.xml', toBytes(`<w:document ${NS}><w:body><w:p>${PICTURE}</w:p><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p></w:body></w:document>`)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function tiff(): Uint8Array {
  const entries = [
    [256, 3, 1, 2], [257, 3, 1, 1], [258, 3, 3, 122],
    [259, 3, 1, 1], [262, 3, 1, 2], [273, 4, 1, 128],
    [277, 3, 1, 3], [278, 4, 1, 1], [279, 4, 1, 6],
  ] as const;
  const bytes = new Uint8Array(134);
  const view = new DataView(bytes.buffer);
  bytes.set([0x49, 0x49, 0x2a, 0]);
  view.setUint32(4, 8, true);
  view.setUint16(8, entries.length, true);
  entries.forEach(([tag, kind, count, value], index) => {
    const at = 10 + index * 12;
    view.setUint16(at, tag, true);
    view.setUint16(at + 2, kind, true);
    view.setUint32(at + 4, count, true);
    view.setUint32(at + 8, value, true);
  });
  for (let index = 0; index < 3; index += 1) view.setUint16(122 + index * 2, 8, true);
  bytes.set([0xff, 0, 0, 0, 0x80, 0xff], 128);
  return bytes;
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  )
);

for (const [extension, mimeType, image] of [
  ['png', 'image/png', PNG],
  ['tif', 'image/tiff', tiff()],
] as const) {
  test(`untouched ${extension} saves are byte-identical with and without media tokens`, async () => {
    const bytes = fixture(image, extension, mimeType);
    const save = async (mediaTokens?: boolean) => {
      const session = await createYrsSession({ clientId: 3 });
      try {
        session.openDocx(bytes, true, mediaTokens === undefined ? undefined : { mediaTokens });
        expect(session.mediaSource('media:1')?.mimeType).toBe('image/png');
        return (await saveYrsDocx(session, { updateModifiedDate: false })).bytes;
      } finally {
        session.destroy();
      }
    };
    const defaultSave = await save();
    expect(unzipContainer(defaultSave)['word/document.xml']).toEqual(
      unzipContainer(bytes)['word/document.xml']
    );
    expect(Buffer.from(await save(false)).equals(defaultSave)).toBe(true);
    expect(Buffer.from(await save(true)).equals(defaultSave)).toBe(true);
  });
}

test('a resident worker opens comment images as a direct open does', async () => {
  const bytes = fixture();
  const direct = await createYrsSession({ clientId: 4 });
  const worker = await createResidentEngineSession();
  try {
    const host = decodeDocxHostJson(worker.openDocx(bytes), bytes);
    const comments = JSON.stringify(host.document.package.document.comments);
    expect(comments).toContain('"src":"data:image/png;base64,');
    expect(comments).not.toContain('media:');
    expect(host.document).toEqual(direct.openDocx(bytes, true).document);
  } finally {
    direct.destroy();
    worker.destroy();
  }
});

test('destroyed sessions return null for cached and uncached media lookups', async () => {
  const session = await createYrsSession({ clientId: 3 });
  try {
    session.openDocx(fixture(), true, { mediaTokens: true });
    expect(session.mediaDataUrl('media:1')).not.toBeNull();
    expect(session.mediaSource('media:1')).not.toBeNull();
    const scope = session.mediaScope();
    session.destroy();
    for (const token of ['media:0', 'media:1', 'media:2']) {
      expect(session.mediaSource(token)).toBeNull();
      expect(session.mediaDataUrl(token)).toBeNull();
    }
    expect(session.mediaScope()).not.toBe(scope);
  } finally {
    session.destroy();
  }
});

for (const mediaTokens of [false, true]) {
  test(`an image resolves through its token and saves untouched (tokens: ${mediaTokens})`, async () => {
    const bytes = fixture();
    const session = await createYrsSession({ clientId: 3 });
    try {
      const host = session.openDocx(bytes, true, { mediaTokens });
      const state = Buffer.from(session.encodeState()).toString('latin1');
      expect(state.includes('data:image/png;base64,')).toBe(!mediaTokens);
      expect(session.mediaSource('media:1')).toEqual({ bytes: PNG, mimeType: 'image/png' });
      expect(session.mediaSource('media:2')).toBeNull();
      const dataUrl = `data:image/png;base64,${Buffer.from(PNG).toString('base64')}`;
      expect(session.mediaDataUrl('media:1')).toBe(dataUrl);
      expect(session.mediaDataUrl(dataUrl)).toBeNull();
      const comments = JSON.stringify(host.document.package.document.comments);
      expect(comments.match(/"src":"[^"]*"/g)?.every((src) => src === `"src":"${dataUrl}"`)).toBe(true);

      const projected = yrsToDocument(session, session.materializeDocx()!);
      const content = JSON.stringify(projected.package.document.content);
      expect(content.match(/"src":"[^"]*"/g)).toEqual([`"src":"${dataUrl}"`]);
      expect(JSON.stringify(projected.package.document.comments)).not.toContain('media:');
      expect(content).toContain('"rId":"rIdImage"');

      const saved = unzipContainer((await saveYrsDocx(session)).bytes);
      const source = unzipContainer(bytes);
      for (const part of ['word/media/picture.png', 'word/media/unused.png']) {
        expect(saved[part]).toEqual(source[part]);
      }
      expect(new TextDecoder().decode(saved['word/document.xml'])).toContain('r:embed="rIdImage"');
    } finally {
      session.destroy();
    }
  });
}
