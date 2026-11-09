import { readFileSync } from 'node:fs';
import { expect, test } from 'bun:test';
import type { DecorationPrimitive, DisplayList, GlyphRunPrimitive } from '../../packages/docx/src/layout/render/displayList';
import { expectPngRule, expectPngTracking, pngPixels } from '../raster-pixels';

type Session = import('./index.js').EditSession;

const raw = require('./index.js') as typeof import('./index.js');
const fontBytes = new Uint8Array(readFileSync(new URL('../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)));
const encoder = new TextEncoder();
const fixture = raw.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/document.xml': encoder.encode('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Engine text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'),
});

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

function pagePng(session: Session, list: string, page: number, scale: number): Uint8Array {
  return session.renderPngTarget(list, `{"page":${page}}`, `{"scale":${scale}}`);
}

function pngDisplayList(
  session: Session,
  families = ['Calibri'],
  font = session.register_measure_font(fontBytes),
): string {
  const fontChains = Object.fromEntries(families.flatMap((family) => [
    [`${family.toLowerCase()}|0|0`, [font]],
    [`${family.toLowerCase()}|1|0`, [font]],
  ]));
  const materialized = JSON.parse(session.materialize_docx()!) as {
    document: { package: { theme: { colorScheme: Record<string, string> } } };
  };
  const layout = session.layout_document_with_regions_json(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{
      sectionId: 'main',
      properties: {
        pageWidth: 4320, pageHeight: 2880,
        marginTop: 300, marginRight: 300, marginBottom: 300, marginLeft: 300,
      },
    }] },
    measurement: {
      fontChains,
      defaults: { fontSize: 11, fontFamily: 'Calibri' },
      authoritativeShaping: true,
    },
    renderEnv: { themeColors: materialized.document.package.theme.colorScheme },
  }));
  return session.build_display_list_json(JSON.stringify({ ...JSON.parse(layout), fontChains }));
}

test('native docx rasterizes the compiled page at two scales', () => {
  const session = new raw.EditSession(401);
  try {
    session.open_docx(fixture, true);
    const list = pngDisplayList(session);
    const first = pagePng(session, list, 0, 1);
    const second = pagePng(session, list, 0, 2);
    for (const png of [first, second]) {
      expectPlainBytes(png);
      expect([...png.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    }
    const one = Buffer.from(first);
    const two = Buffer.from(second);
    expect(two.readUInt32BE(16)).toBe(one.readUInt32BE(16) * 2);
    expect(two.readUInt32BE(20)).toBe(one.readUInt32BE(20) * 2);
    expect(pngPixels(first).some((value, index) => index % 4 < 3 && value < 200)).toBe(true);
    expect(pngPixels(second).some((value, index) => index % 4 < 3 && value < 200)).toBe(true);
  } finally {
    session.free();
  }
});

test('native docx draws a run scale of 0 at full width, as Word does', () => {
  const scaled = (properties: string) => {
    const session = new raw.EditSession(405);
    try {
      session.open_docx(raw.rezip_docx({
        '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
        '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
        'word/document.xml': encoder.encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r>${properties}<w:t>Engine text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`),
      }), true);
      return Buffer.from(pagePng(session, pngDisplayList(session), 0, 1));
    } finally {
      session.free();
    }
  };
  const plain = scaled('');
  expect(scaled('<w:rPr><w:w w:val="0"/></w:rPr>').equals(plain)).toBe(true);
  expect(scaled('<w:rPr><w:w w:val="50"/></w:rPr>').equals(plain)).toBe(false);
});

test('native docx PNG refuses invalid input, missing fonts and oversized surfaces', () => {
  const session = new raw.EditSession(402);
  try {
    session.open_docx(fixture, true);
    const list = pngDisplayList(session);
    for (const call of [
      () => pagePng(session, list, 9999, 1),
      () => pagePng(session, list, 0, 0),
      () => pagePng(session, list, 0, NaN),
      () => pagePng(session, list, 0, -1),
      () => pagePng(session, list, 0, Infinity),
      () => pagePng(session, '{', 0, 1),
      () => pagePng(session, '{}', 0, 1),
      () => pagePng(session, JSON.stringify({ pages: [{ pageIndex: 0, width: 0, height: 10, primitives: [] }] }), 0, 1),
    ]) {
      const error = caughtValue(call);
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).message).toStartWith('png-invalid:');
      expect(error).not.toHaveProperty('code');
    }
    for (const call of [
      () => session.renderPngTarget(list, '{"page":0}', '{"scale":null}'),
      () => session.renderPngTarget(list, '{"page":0}', '{"scale":"2"}'),
      () => session.renderPngTarget(list, '{"page":0}', '{"maxWidth":10}'),
      () => session.renderPngTarget(list, '{"page":0,"extra":1}', '{}'),
      () => session.renderPngTarget(list, '{"slide":0}', '{}'),
      () => session.renderPngTarget(list, '{"page":-1}', '{}'),
      () => session.renderPngTarget(list, '{"page":0.5}', '{}'),
      () => session.renderPngTarget(list, '{}', '{}'),
      () => session.renderPngTarget(list, '{"page":0}', '['),
      () => session.renderPngTarget(list, '{"page":18446744073709551615}', '{}'),
      () => pagePng(session, JSON.stringify({ pages: [{ ...JSON.parse(list).pages[0], unbuilt: true }] }), 0, 1),
    ]) {
      const error = caughtValue(call);
      expect(error).toBeInstanceOf(Error);
      expect((error as Error).message).toStartWith('png-invalid:');
      expect(error).not.toHaveProperty('code');
    }
    expect(session.renderPngTarget(list, '{"page":0}', '{}')).toEqual(pagePng(session, list, 0, 1));
    for (const scale of [1e9, Number.MAX_VALUE]) {
      expect(() => pagePng(session, list, 0, scale)).toThrow('png-too-large:');
    }
    const area = JSON.stringify({ pages: [{ pageIndex: 0, width: 5000, height: 5000, primitives: [] }] });
    expect(() => pagePng(session, area, 0, 1)).toThrow('png-too-large:');
    const stale = JSON.parse(list);
    const glyph = stale.pages[0].primitives.find((primitive: { kind: string }) => primitive.kind === 'glyphRun');
    expect(glyph).toBeDefined();
    glyph.fontId = 0x7ffffffe;
    expect(() => pagePng(session, JSON.stringify(stale), 0, 1)).toThrow('png-font-unavailable:');
    const other = new raw.EditSession(404);
    try {
      other.open_docx(fixture, true);
      expect(() => pagePng(other, list, 0, 1)).toThrow('png-font-unavailable:');
      expect(pagePng(session, list, 0, 1).length).toBeGreaterThan(8);
    } finally {
      other.free();
    }
  } finally {
    session.free();
  }
});

test('native docx PNG paints theme accent text red', () => {
  const encoder = new TextEncoder();
  const themed = raw.rezip_docx({
    '[Content_Types].xml': encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/_rels/document.xml.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>'),
    'word/theme/theme1.xml': encoder.encode('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Red"><a:themeElements><a:clrScheme name="Red"><a:accent1><a:srgbClr val="FF0000"/></a:accent1></a:clrScheme><a:fontScheme name="Default"><a:majorFont><a:latin typeface="Calibri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Default"/></a:themeElements></a:theme>'),
    'word/document.xml': encoder.encode('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/><w:sz w:val="96"/><w:color w:themeColor="accent1"/></w:rPr><w:t>RED</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'),
  });
  const session = new raw.EditSession(403);
  try {
    session.open_docx(themed, true);
    const pixels = pngPixels(pagePng(session, pngDisplayList(session), 0, 1));
    let red = false;
    for (let index = 0; index < pixels.length; index += 4) {
      if (pixels[index] > 200 && pixels[index + 1] < 50 && pixels[index + 2] < 50 && pixels[index + 3] === 255) red = true;
    }
    expect(red).toBe(true);
  } finally {
    session.free();
  }
});

test('native docx PNG uses display list font chains for text primitives', () => {
  const session = new raw.EditSession(405);
  try {
    session.open_docx(fixture, true);
    const font = session.register_measure_font(fontBytes);
    const pages = [{ pageIndex: 0, width: 160, height: 60, primitives: [{
      kind: 'text', text: 'PNG', x: 5, baselineY: 30, width: 100,
      font: '24px Calibri', color: '#000000',
    }] }];
    const list = JSON.stringify({ pages, fontChains: { 'calibri|0|0': [font] } });
    expect(pngPixels(pagePng(session, list, 0, 1)).some((value, index) => index % 4 < 3 && value < 200)).toBe(true);
    expect(() => pagePng(session, JSON.stringify({ pages }), 0, 1)).toThrow('png-font-unavailable:');
  } finally {
    session.free();
  }
});

test('native docx PNG resolves media tokens in every page region and refuses a picture it cannot draw', () => {
  const parts = raw.unzip_docx(fixture) as Record<string, Uint8Array>;
  const blue = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYPj/HwADAgH/5ncLrgAAAABJRU5ErkJggg==';
  parts['word/media/image1.png'] = new Uint8Array(Buffer.from(blue, 'base64'));
  parts['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(parts['[Content_Types].xml'])
    .replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'));
  const session = new raw.EditSession(406);
  try {
    session.set_media_tokens(true);
    session.open_docx(raw.rezip_docx(parts), true);
    expect(session.media_bytes('media:0')).toEqual(parts['word/media/image1.png']);
    const image = (x: number, relId = 'media:0') => ({ kind: 'image', relId, x, y: 10, w: 10, h: 10 });
    const list = JSON.stringify({ pages: [{
      pageIndex: 0, width: 160, height: 60,
      primitives: [image(10), image(130, `data:image/png;base64,${blue}`), {
        kind: 'shape', x: 145, y: 10, w: 10, h: 10,
        geometryPath: [{ type: 'move', x: 145, y: 10 }, { type: 'line', x: 155, y: 10 }, { type: 'line', x: 155, y: 20 }, { type: 'line', x: 145, y: 20 }, { type: 'close' }],
        fillPaint: { kind: 'picture', pictureSrc: 'media:0', pictureFillMode: 'stretch' },
      }],
      header: { rId: 'header1', kind: 'header', y: 0, height: 30, primitives: [image(30)] },
      footer: { rId: 'footer1', kind: 'footer', y: 30, height: 30, primitives: [image(50)] },
      noteAreas: [{ primitives: [image(70)] }, { kind: 'endnote', primitives: [image(90)] }],
    }] });
    for (const scale of [1, 2]) {
      const png = pagePng(session, list, 0, scale);
      const width = Buffer.from(png).readUInt32BE(16);
      const pixels = pngPixels(png);
      const pixel = (x: number) => {
        const offset = (15 * scale * width + x * scale) * 4;
        return [...pixels.subarray(offset, offset + 4)];
      };
      for (const x of [15, 35, 55, 75, 95, 135, 150]) expect(pixel(x)).toEqual([0, 0, 255, 255]);
      expect(pixel(115)).toEqual([255, 255, 255, 255]);
    }
    const dropped = JSON.stringify({ pages: [{ pageIndex: 0, width: 160, height: 60, primitives: [image(10), image(110, 'media:999')] }] });
    expect(() => pagePng(session, dropped, 0, 1)).toThrow('png-picture-unavailable: 1 picture(s) on this page could not be drawn');
  } finally {
    session.free();
  }
});

test('native docx PNG draws an EMF picture that carries text from the metafile itself', () => {
  const parts = raw.unzip_docx(fixture) as Record<string, Uint8Array>;
  parts['word/media/image1.emf'] = new Uint8Array(readFileSync(new URL('../../crates/ooxml-metafile/tests/fixtures/libreoffice-one-line.emf', import.meta.url)));
  parts['[Content_Types].xml'] = new TextEncoder().encode(new TextDecoder().decode(parts['[Content_Types].xml'])
    .replace('</Types>', '<Default Extension="emf" ContentType="image/x-emf"/></Types>'));
  const session = new raw.EditSession(407);
  try {
    session.set_media_tokens(true);
    session.open_docx(raw.rezip_docx(parts), true);
    const list = JSON.stringify({ pages: [{ pageIndex: 0, width: 200, height: 60,
      primitives: [{ kind: 'image', relId: 'media:0', x: 0, y: 0, w: 200, h: 60 }] }] });
    const png = pagePng(session, list, 0, 1);
    const pixels = pngPixels(png);
    let painted = 0;
    for (let at = 0; at < pixels.length; at += 4) if (pixels[at] < 200 || pixels[at + 1] < 200 || pixels[at + 2] < 200) painted++;
    expect(painted).toBeGreaterThan(20);
  } finally {
    session.free();
  }
});

function docxTextRaster(family: string, properties = '') {
  const parts = raw.unzip_docx(fixture) as Record<string, Uint8Array>;
  parts['word/document.xml'] = new TextEncoder().encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/><w:sz w:val="24"/>${properties}</w:rPr><w:t>TTTT</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const session = new raw.EditSession(410);
  try {
    session.open_docx(raw.rezip_docx(parts), true);
    const face = session.register_measure_font(fontBytes);
    const json = pngDisplayList(session, ['Calibri', 'Liberation Sans', 'Unregistered Raster Face'], face);
    const list = JSON.parse(json) as DisplayList;
    const runs = list.pages[0].primitives.filter((primitive): primitive is GlyphRunPrimitive => primitive.kind === 'glyphRun');
    expect(runs.length).toBeGreaterThan(0);
    for (const part of runs) {
      expect(part.fontId).toBe(face);
      expect(part.size).toBe(16);
    }
    const run = { ...runs[0], text: runs.map((part) => part.text).join(''), glyphs: runs.flatMap((part) => part.glyphs) };
    expect(run.text).toBe('TTTT');
    expect(run.glyphs.length).toBe(4);
    return { session, json, list, run };
  } catch (error) {
    session.free();
    throw error;
  }
}

test.each([1, 2])('native docx PNG preserves the layout fallback face at scale %s', (scale) => {
  const fallback = docxTextRaster('Unregistered Raster Face');
  const direct = docxTextRaster('Liberation Sans');
  try {
    expect(fallback.run.fontId).toBe(direct.run.fontId);
    expect(fallback.run.glyphs).toEqual(direct.run.glyphs);
    expect(fallback.run.size).toBe(direct.run.size);
    const actual = pagePng(fallback.session, fallback.json, 0, scale);
    const expected = pagePng(direct.session, direct.json, 0, scale);
    expect(pngPixels(actual)).toEqual(pngPixels(expected));
    expect(pngPixels(actual).some((value, index) => index % 4 < 3 && value < 200)).toBe(true);
  } finally {
    fallback.session.free();
    direct.session.free();
  }
});

test.each([1, 2])('native docx PNG preserves positioned letter spacing at scale %s', (scale) => {
  const spaced = docxTextRaster('Liberation Sans', '<w:spacing w:val="200"/>');
  const plain = docxTextRaster('Liberation Sans');
  try {
    const glyphs = (run: GlyphRunPrimitive) => run.glyphs.map((glyph) => {
      expect(glyph.advance).toBeDefined();
      return { x: glyph.x, advance: glyph.advance! };
    });
    const positioned = glyphs(spaced.run);
    const last = positioned[positioned.length - 1];
    const inked = [...positioned.slice(0, -1), { ...last, advance: last.advance - 200 / 15 }];
    expectPngTracking(
      pagePng(spaced.session, spaced.json, 0, scale),
      pagePng(plain.session, plain.json, 0, scale),
      inked, glyphs(plain.run), 200 / 15, scale,
    );
  } finally {
    spaced.session.free();
    plain.session.free();
  }
});

for (const [decoration, properties] of [
  ['underline', '<w:u w:val="single"/>'],
  ['strike', '<w:strike/>'],
] as const) {
  test.each([1, 2])(`native docx PNG paints display list ${decoration} at scale %s`, (scale) => {
    const decorated = docxTextRaster('Liberation Sans', properties);
    const plain = docxTextRaster('Liberation Sans');
    try {
      expect(decorated.run.glyphs).toEqual(plain.run.glyphs);
      const rules = decorated.list.pages[0].primitives.filter((primitive): primitive is DecorationPrimitive =>
        primitive.kind === 'decoration' && primitive.deco === decoration);
      expect(rules.length).toBeGreaterThan(0);
      expect(plain.list.pages[0].primitives.some((primitive) => primitive.kind === 'decoration')).toBe(false);
      rules.sort((left, right) => left.x - right.x);
      for (let index = 1; index < rules.length; index++) {
        expect(rules[index].y).toBe(rules[0].y);
        expect(rules[index].h).toBe(rules[0].h);
        expect(Math.abs(rules[index].x - rules[index - 1].x - rules[index - 1].w)).toBeLessThan(0.02);
      }
      const end = rules[rules.length - 1];
      const rule = { ...rules[0], w: end.x + end.w - rules[0].x };
      const first = decorated.run.glyphs[0];
      const last = decorated.run.glyphs[decorated.run.glyphs.length - 1];
      expect(Math.abs(rule.x - first.x)).toBeLessThan(0.02);
      expect(Math.abs(rule.w - (last.x + last.advance! - first.x))).toBeLessThan(0.02);
      expectPngRule(
        pagePng(decorated.session, decorated.json, 0, scale),
        pagePng(plain.session, plain.json, 0, scale),
        rule, scale,
      );
    } finally {
      decorated.session.free();
      plain.session.free();
    }
  });
}
