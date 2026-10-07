import { beforeAll, expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { initWasm, inspectPresentation } from '../index';
import { parsePptxJson, parsePptxJsonWithoutMedia } from './generated/pptx_wasm.js';

const media = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><rect width="1" height="1"/></svg>');
let deck: Uint8Array;

beforeAll(async () => {
  await initWasm(await readFile(resolve(import.meta.dir, 'generated/pptx_wasm_bg.wasm')));
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="svg" ContentType="image/svg+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="914400" cy="914400"/></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld></p:sld>');
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.svg"/></Relationships>');
  zip.file('ppt/media/image1.svg', media);
  deck = await zip.generateAsync({ type: 'uint8array' });
});

test('inspection omits media bytes while preserving media metadata and deck structure', () => {
  const full = inspectPresentation(deck) as { media: unknown[] };
  const metadata = inspectPresentation(deck, { includeMedia: false }) as { media: unknown[] };
  expect(metadata.media[0]).not.toHaveProperty('bytes');
  expect(metadata.media).toEqual([{
    partPath: 'ppt/media/image1.svg', contentType: 'image/svg+xml', byteLength: media.length,
  }]);
  expect(metadata).toEqual({ ...full, media: metadata.media });
  expect(metadata).toMatchObject({
    slides: [{ shapes: [{ kind: 'picture', mediaPartPath: 'ppt/media/image1.svg' }] }],
  });
  expect(JSON.stringify(metadata)).not.toContain(Buffer.from(media).toString('base64'));
});

test('inspection includes unchanged media bytes by default', () => {
  const full = inspectPresentation(deck) as { media: unknown[] };
  expect(full.media).toEqual([{
    partPath: 'ppt/media/image1.svg', contentType: 'image/svg+xml',
    bytes: Buffer.from(media).toString('base64'),
  }]);
  const legacy = JSON.stringify(JSON.parse(parsePptxJson(deck)));
  expect(JSON.stringify(full)).toBe(legacy);
  expect(JSON.stringify(inspectPresentation(deck, {}))).toBe(legacy);
  expect(JSON.stringify(inspectPresentation(deck, { includeMedia: true }))).toBe(legacy);
});

test('inspection without media preserves the raw JSON outside the media array', () => {
  const full = parsePptxJson(deck);
  const metadata = parsePptxJsonWithoutMedia(deck);
  const fullMedia = JSON.parse(full).media;
  const summaries = [{
    partPath: 'ppt/media/image1.svg', contentType: 'image/svg+xml', byteLength: media.length,
  }];
  expect(fullMedia).toHaveLength(1);
  expect(full).toContain(`"media":${JSON.stringify(fullMedia)}`);
  expect(metadata).toBe(full.replace(
    `"media":${JSON.stringify(fullMedia)}`,
    `"media":${JSON.stringify(summaries)}`,
  ));
});
