import { readFileSync } from 'node:fs';
import { expect, test } from 'bun:test';
import JSZip from 'jszip';
import type { SlideDisplayList } from '../../packages/pptx/src/types';
import { expectPngRule, expectPngTracking, pngPixels } from '../raster-pixels';

const raw = require('./index.js') as typeof import('./index.js');
const fixture = readFileSync(new URL('../../apps/demo/public/betteroffice-demo.pptx', import.meta.url));
const bytes = () => new Uint8Array(fixture);

/** Raw `renderPngTarget` with the scale written as-is, so NaN and Infinity reach the binding as bad JSON. */
function slidePng(renderer: { renderPngTarget(document: unknown, target: string, options: string): Uint8Array }, document: unknown, slide: number, scale: number): Uint8Array {
  return renderer.renderPngTarget(document, `{"slide":${slide}}`, `{"scale":${scale}}`);
}

function expectPlainBytes(value: Uint8Array) {
  expect(value).toBeInstanceOf(Uint8Array);
  expect(Buffer.isBuffer(value)).toBe(false);
  expect(Object.getPrototypeOf(value)).toBe(Uint8Array.prototype);
  expect(value.byteOffset).toBe(0);
  expect(value.buffer.byteLength).toBe(value.byteLength);
}

function caughtValue(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  throw new Error('Expected the engine call to throw');
}

test('native pptx rasterizes a scoped slide at two scales', () => {
  const document = raw.PptxDocument.openCollaborative(bytes(), 401);
  const renderer = new raw.PptxRenderer();
  try {
    const font = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
    renderer.registerFont('Arial', false, false, font);
    const first = slidePng(renderer, document, 0, 1);
    const second = slidePng(renderer, document, 0, 2);
    for (const png of [first, second]) {
      expectPlainBytes(png);
      expect([...png.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    }
    const one = Buffer.from(first);
    const two = Buffer.from(second);
    expect(two.readUInt32BE(16)).toBe(one.readUInt32BE(16) * 2);
    expect(two.readUInt32BE(20)).toBe(one.readUInt32BE(20) * 2);
    for (const call of [
      () => slidePng(renderer, document, 9999, 1),
      () => slidePng(renderer, document, 0, 0),
      () => slidePng(renderer, document, 0, NaN),
      () => slidePng(renderer, document, 0, -1),
      () => slidePng(renderer, document, 0, Infinity),
    ]) {
      const error = caughtValue(call);
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).message).toStartWith('png-invalid:');
      expect(error).not.toHaveProperty('code');
    }
    for (const scale of [10, 1e9, Number.MAX_VALUE]) {
      expect(() => slidePng(renderer, document, 0, scale)).toThrow('png-too-large:');
    }
    expect(slidePng(renderer, document, 0, 1).length).toBeGreaterThan(8);
  } finally {
    renderer.free();
    document.free();
  }
});

test('native pptx PNG rejects a slide with a picture it cannot draw', async () => {
  const zip = await JSZip.loadAsync(bytes());
  zip.file('ppt/media/betteroffice-mark.png', new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]));
  const document = raw.PptxDocument.openCollaborative(await zip.generateAsync({ type: 'uint8array' }), 411);
  const renderer = new raw.PptxRenderer();
  try {
    const font = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
    renderer.registerFont('Arial', false, false, font);
    expect(() => slidePng(renderer, document, 0, 1)).toThrow('png-picture-unavailable: 1 picture(s) on this slide could not be drawn');
    expect(slidePng(renderer, document, 1, 1).length).toBeGreaterThan(8);
  } finally {
    renderer.free();
    document.free();
  }
});

test('native pptx PNG draws a slide whose picture has no image to draw', async () => {
  const zip = await JSZip.loadAsync(bytes());
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
  const emptied = slide.replace('<a:blip r:embed="rId2"/>', '<a:blip/>');
  expect(emptied).not.toBe(slide);
  zip.file('ppt/slides/slide1.xml', emptied);
  const document = raw.PptxDocument.openCollaborative(await zip.generateAsync({ type: 'uint8array' }), 412);
  const renderer = new raw.PptxRenderer();
  try {
    const font = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
    renderer.registerFont('Arial', false, false, font);
    expect(slidePng(renderer, document, 0, 1).length).toBeGreaterThan(8);
  } finally {
    renderer.free();
    document.free();
  }
});

test('native pptx PNG requires registered fonts for slide text', () => {
  const document = raw.PptxDocument.openCollaborative(bytes(), 402);
  const renderer = new raw.PptxRenderer();
  try {
    const slideId = (JSON.parse(document.snapshotJson()) as { slides: { id: string }[] }).slides[0].id;
    document.addTextBoxJson(JSON.stringify({ slideId, draft: {
      name: 'PNG text', rect: { x: 100000, y: 100000, width: 3000000, height: 1000000 },
      text: 'Registered face required', style: { fontFamily: 'Arial', fontSizePt: 18, color: '#000000' },
    } }));
    const error = caughtValue(() => slidePng(renderer, document, 0, 1));
    expect(error).toBeInstanceOf(Error);
    expect((error as Error).message).toStartWith('png-font-unavailable:');
    expect(error).not.toHaveProperty('code');
  } finally {
    renderer.free();
    document.free();
  }
});

async function pptxTextRaster(family: string, properties = '') {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="3048000" cy="1524000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Raster text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="190500" y="190500"/><a:ext cx="2667000" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="l" marL="0" indent="0"/><a:r><a:rPr sz="1200" ${properties}><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="${family}"/></a:rPr><a:t>TTTT</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  const document = raw.PptxDocument.openCollaborative(await zip.generateAsync({ type: 'uint8array' }), 410);
  const renderer = new raw.PptxRenderer();
  try {
    const font = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
    const face = renderer.registerFont('Liberation Sans', false, false, font);
    const list = JSON.parse(renderer.layoutSlideJson(document, 0)) as SlideDisplayList;
    const boxes = list.primitives.filter((primitive) => primitive.kind === 'textBox');
    expect(boxes.length).toBe(1);
    const lines = boxes.flatMap((box) => box.kind === 'textBox' ? box.lines : []);
    expect(lines.length).toBe(1);
    expect(lines[0].runs.length).toBe(1);
    const run = lines[0].runs[0];
    expect(run.text).toBe('TTTT');
    expect(run.fontId).toBe(face);
    expect(run.fontSizePx).toBe(16);
    expect(run.glyphs.length).toBe(4);
    return { document, renderer, run, line: lines[0] };
  } catch (error) {
    renderer.free();
    document.free();
    throw error;
  }
}

test.each([1, 2])('native pptx PNG preserves the layout fallback face at scale %s', async (scale) => {
  const fallback = await pptxTextRaster('Unregistered Raster Face');
  const direct = await pptxTextRaster('Liberation Sans');
  try {
    expect(fallback.run.fontId).toBe(direct.run.fontId);
    expect(fallback.run.glyphs).toEqual(direct.run.glyphs);
    expect(fallback.run.fontSizePx).toBe(direct.run.fontSizePx);
    const actual = slidePng(fallback.renderer, fallback.document, 0, scale);
    const expected = slidePng(direct.renderer, direct.document, 0, scale);
    expect(pngPixels(actual)).toEqual(pngPixels(expected));
    expect(pngPixels(actual).some((value, index) => index % 4 < 3 && value < 200)).toBe(true);
  } finally {
    fallback.renderer.free();
    fallback.document.free();
    direct.renderer.free();
    direct.document.free();
  }
});

test.each([1, 2])('native pptx PNG preserves positioned letter spacing at scale %s', async (scale) => {
  const spaced = await pptxTextRaster('Liberation Sans', 'spc="1000"');
  const plain = await pptxTextRaster('Liberation Sans');
  try {
    expect(spaced.run.letterSpacingPx).toBeCloseTo(1000 / 75, 4);
    expectPngTracking(
      slidePng(spaced.renderer, spaced.document, 0, scale),
      slidePng(plain.renderer, plain.document, 0, scale),
      spaced.run.glyphs, plain.run.glyphs, 1000 / 75, scale,
    );
  } finally {
    spaced.renderer.free();
    spaced.document.free();
    plain.renderer.free();
    plain.document.free();
  }
});

for (const [decoration, properties] of [
  ['underline', 'u="sng"'],
  ['strike', 'strike="sngStrike"'],
] as const) {
  test.each([1, 2])(`native pptx PNG paints display list ${decoration} at scale %s`, async (scale) => {
    const decorated = await pptxTextRaster('Liberation Sans', properties);
    const plain = await pptxTextRaster('Liberation Sans');
    try {
      const { run, line } = decorated;
      expect(run.glyphs).toEqual(plain.run.glyphs);
      expect(plain.run.underline).toBe(false);
      expect(plain.run.strike ?? 'noStrike').toBe('noStrike');
      if (decoration === 'underline') expect(run.underline).toBe(true);
      else expect(run.strike).toBe('sngStrike');
      const first = run.glyphs[0];
      const last = run.glyphs[run.glyphs.length - 1];
      expect(Math.abs(run.x - first.x)).toBeLessThan(0.02);
      expect(Math.abs(run.width - (last.x + last.advance - first.x))).toBeLessThan(0.02);
      const baseline = line.baseline - (run.baselineOffsetPx ?? 0);
      const rule = {
        x: run.x,
        y: baseline + run.fontSizePx * (decoration === 'underline' ? 0.08 : -0.3),
        w: run.width,
        h: Math.max(1, run.fontSizePx * 0.05),
      };
      expectPngRule(
        slidePng(decorated.renderer, decorated.document, 0, scale),
        slidePng(plain.renderer, plain.document, 0, scale),
        rule, scale,
      );
    } finally {
      decorated.renderer.free();
      decorated.document.free();
      plain.renderer.free();
      plain.document.free();
    }
  });
}

test('default fonts leave the other styles of a registered family to that family', async () => {
  const assets = new URL('../../packages/fonts-latin/assets', import.meta.url).pathname;
  raw.configureFontPacks(assets, undefined, undefined);
  const { document, renderer } = await pptxTextRaster('Liberation Sans', 'b="1"');
  try {
    const before = slidePng(renderer, document, 0, 1);
    renderer.registerDefaultFonts(document, 0);
    expect(slidePng(renderer, document, 0, 1)).toEqual(before);
  } finally {
    renderer.free();
    document.free();
  }
});
