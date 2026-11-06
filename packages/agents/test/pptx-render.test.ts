import { beforeAll, expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { createFontProvider } from '@betteroffice/fonts';
import { initWasm, openPresentation, type SlideSnapshot } from '@betteroffice/pptx';
import { pptxFontFaces, pptxThemeFonts, renderPptxSlide } from '../src/pptx-render';
import { initialize } from '../src/prototype';
import { attachPptx } from '../src/pptx';
import { createCanvas } from '@napi-rs/canvas';
import { presentationFixture } from './pptx-fixture';

beforeAll(() => initialize('pptx', initWasm));

for (const failPrimary of [false, true]) {
  test(`PPTX preview skips rejected coverage with ${failPrimary ? 'a rejected' : 'a loaded'} primary face`, async () => {
    const zip = await JSZip.loadAsync(await presentationFixture());
    const path = 'ppt/slides/slide1.xml';
    zip.file(path, (await zip.file(path)!.async('string')).replaceAll('</a:rPr>', '<a:latin typeface="Segoe UI"/></a:rPr>'));
    const handle = openPresentation(await zip.generateAsync({ type: 'uint8array' }));
    const render = spyOn(handle, 'renderPng').mockRejectedValue(Object.assign(new Error('canvas unavailable'), { code: 'canvas-unavailable' }));
    const provider = createFontProvider();
    const lastResorts: string[] = [];
    let rejectedCoverage = 0;
    try {
      const result = await renderPptxSlide(handle, 1, 1, {
        ...provider,
        resolve: (family, bold, italic) => failPrimary && family === 'Segoe UI'
          ? async () => { throw new Error('primary unavailable'); }
          : provider.resolve(family, bold, italic),
        resolveLastResort: (family, bold, italic, office) => {
          lastResorts.push(family);
          return provider.resolveLastResort(family, bold, italic, office);
        },
        resolveGlyphFallback: (family, bold, italic, office) => provider.resolveGlyphFallback!(family, bold, italic, office)
          ? async () => { rejectedCoverage++; throw new Error('coverage unavailable'); }
          : undefined,
      });
      expect(result.png.length).toBeGreaterThan(0);
      expect(rejectedCoverage).toBeGreaterThan(0);
      expect(lastResorts.includes('Segoe UI')).toBe(failPrimary);
    } finally { render.mockRestore(); handle.dispose(); }
  });
}

test('PPTX save faces follow the theme for runs that name no face', async () => {
  const zip = new JSZip();
  zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="s"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>');
  const theme = await pptxThemeFonts(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  expect(theme).toEqual({ major: ['Calibri Light'], minor: ['Calibri'] });
  expect(await pptxThemeFonts(await zip.generateAsync({ type: 'uint8array' }))).toEqual(theme);
  expect(await pptxThemeFonts(new Uint8Array(64))).toEqual({ major: [], minor: [] });
  const prefixed = new JSZip();
  prefixed.file('ppt/theme/theme1.xml', '<d:theme xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><d:themeElements><d:fontScheme name="s"><d:majorFont><d:latin typeface="Georgia"/></d:majorFont><d:minorFont><d:latin typeface="Verdana"/></d:minorFont></d:fontScheme></d:themeElements></d:theme>');
  prefixed.file('ppt/theme/theme2.xml', '<theme xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><themeElements><fontScheme name="s"><majorFont><latin typeface="Tahoma"/></majorFont><minorFont><latin typeface="Verdana"/></minorFont></fontScheme></themeElements></theme>');
  expect(await pptxThemeFonts(await prefixed.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))).toEqual({ major: ['Georgia', 'Tahoma'], minor: ['Verdana'] });
  const run = (fontFamily: string | undefined, bold = false) => ({ text: 'x', style: { fontFamily, bold } });
  const slide = { shapes: [{ textStories: [{ paragraphs: [{ runs: [run(undefined, true), run('+mj-lt'), run('Georgia')] }] }], children: [] }] } as unknown as SlideSnapshot;
  expect(pptxFontFaces([slide], theme)).toEqual([
    { family: 'Arial', bold: false, italic: false },
    { family: 'Calibri Light', bold: true, italic: false },
    { family: 'Calibri', bold: true, italic: false },
    { family: 'Calibri Light', bold: false, italic: false },
    { family: 'Georgia', bold: false, italic: false },
  ]);
  expect(pptxFontFaces([slide])).toEqual([
    { family: 'Arial', bold: false, italic: false },
    { family: 'Arial', bold: true, italic: false },
    { family: 'Georgia', bold: false, italic: false },
  ]);
});

test('PPTX theme fonts include custom theme part names', async () => {
  const zip = new JSZip();
  zip.file('ppt/theme/corporate.xml', '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="s"><a:majorFont><a:latin typeface="Times New Roman"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>');
  zip.file('ppt/theme/_rels/corporate.xml', '<a:majorFont><a:latin typeface="Ignored"/></a:majorFont>');
  expect(await pptxThemeFonts(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))).toEqual({ major: ['Times New Roman'], minor: ['Arial'] });
});

test('PPTX theme fonts accept single-quoted typefaces', async () => {
  const zip = new JSZip();
  zip.file('ppt/theme/theme1.xml', `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="s"><a:majorFont><a:latin typeface='Times New Roman'/></a:majorFont><a:minorFont><a:latin typeface='Arial'/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`);
  expect(await pptxThemeFonts(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))).toEqual({ major: ['Times New Roman'], minor: ['Arial'] });
});

test('PPTX theme fonts accept whitespace around typeface equals', async () => {
  const zip = new JSZip();
  zip.file('ppt/theme/theme1.xml', `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="s"><a:majorFont><a:latin xtypeface='Decoy' typeface = 'Times New Roman'/></a:majorFont><a:minorFont><a:latin typeface =\n"Arial"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`);
  expect(await pptxThemeFonts(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))).toEqual({ major: ['Times New Roman'], minor: ['Arial'] });
});

test('PPTX theme fonts ignore XML comment decoys', async () => {
  const zip = new JSZip();
  zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><!--<a:majorFont><a:latin typeface="Decoy"/></a:majorFont>--><a:themeElements><a:fontScheme name="s"><a:majorFont><a:latin typeface="Times New Roman"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>');
  expect(await pptxThemeFonts(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))).toEqual({ major: ['Times New Roman'], minor: ['Arial'] });
});

test('PPTX theme fonts scan large repetitive comments within two seconds', async () => {
  const zip = new JSZip();
  zip.file('ppt/theme/theme1.xml', `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><!--${'<a:majorFont '.repeat(200000)}--><a:themeElements><a:fontScheme name="s"><a:majorFont><a:latin typeface="Times New Roman"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`);
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const start = performance.now();
  const theme = await pptxThemeFonts(bytes);
  const elapsed = performance.now() - start;
  expect(theme).toEqual({ major: ['Times New Roman'], minor: ['Arial'] });
  expect(elapsed).toBeLessThan(2000);
});

test('PPTX save faces expand theme families once per style, not once per run', () => {
  const theme = { major: Array.from({ length: 4000 }, (_, index) => `Major ${index}`), minor: Array.from({ length: 4000 }, (_, index) => `Minor ${index}`) };
  const runs = Array.from({ length: 50000 }, (_, index) => ({ text: 'x', style: { bold: index % 2 === 0 } }));
  const slides = [{ shapes: [{ textStories: [{ paragraphs: [{ runs }] }], children: [] }] }] as unknown as Parameters<typeof pptxFontFaces>[0];
  const started = performance.now();
  const faces = pptxFontFaces(slides, theme);
  const limited = pptxFontFaces(slides, theme, 192);
  expect(performance.now() - started).toBeLessThan(2000);
  expect(faces).toHaveLength(1 + 2 * 8000);
  expect(faces.slice(0, 3)).toEqual([{ family: 'Arial', bold: false, italic: false }, { family: 'Major 0', bold: true, italic: false }, { family: 'Major 1', bold: true, italic: false }]);
  expect(limited).toEqual(faces.slice(0, 192));
});

test('PPTX theme fonts read only the font scheme, not font elements in extensions', async () => {
  const zip = new JSZip();
  const decoys = Array.from({ length: 20000 }, (_, index) => `<a:majorFont><a:latin typeface="Decoy ${index}"/></a:majorFont><a:minorFont><a:latin typeface="Other ${index}"/></a:minorFont>`).join('');
  zip.file('ppt/theme/theme1.xml', `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="s"><a:majorFont><a:latin typeface="Times New Roman"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme></a:themeElements><a:extLst><a:ext uri="{00000000-0000-0000-0000-000000000000}">${decoys}</a:ext></a:extLst></a:theme>`);
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const started = performance.now();
  expect(await pptxThemeFonts(bytes)).toEqual({ major: ['Times New Roman'], minor: ['Arial'] });
  expect(performance.now() - started).toBeLessThan(2000);
});

test('PPTX previews measure runs without a face in the theme fonts the export uses', async () => {
  const handle = openPresentation(await presentationFixture());
  const registered: string[] = [];
  const tracked = { ...handle, renderPng: async () => { throw Object.assign(new Error('canvas unavailable'), { code: 'canvas-unavailable' }); }, registerFont: (face: Parameters<typeof handle.registerFont>[0]) => { registered.push(face.family); return handle.registerFont(face); } };
  try {
    const theme = { major: ['Georgia'], minor: ['Verdana'] };
    await renderPptxSlide(tracked, 1, 1, createFontProvider(), theme);
    const runs = handle.snapshot().slides[0].shapes.flatMap(shape => shape.textStories).flatMap(story => story.paragraphs).flatMap(paragraph => paragraph.runs);
    expect(runs.some(run => !run.style.fontFamily)).toBe(true);
    expect(registered).toEqual(expect.arrayContaining(['Georgia', 'Verdana']));
  } finally { handle.dispose(); }
});

test('PPTX previews use the engine PNG where it can draw', async () => {
  const handle = openPresentation(await presentationFixture());
  const deck = attachPptx(handle);
  const png = new Uint8Array(await createCanvas(12, 7).encode('png'));
  const render = spyOn(handle, 'renderPng').mockResolvedValue(png);
  const layout = spyOn(handle, 'layoutSlide');
  try {
    const slide = handle.snapshot().slides[0]!;
    for (const preview of [await deck.preview(slide.id), await deck.render(1)]) {
      expect(preview.png).toBe(png);
      expect(preview).toMatchObject({ page: 1, pageCount: handle.snapshot().slides.length, width: 12, height: 7 });
    }
    expect(render).toHaveBeenCalledTimes(2);
    expect(render).toHaveBeenCalledWith({ slide: slide.id }, { scale: 1 });
    expect(layout).not.toHaveBeenCalled();
  } finally { layout.mockRestore(); render.mockRestore(); deck.close(); handle.dispose(); }
});

test('PPTX previews fall back to canvas when the engine cannot draw', async () => {
  const handle = openPresentation(await presentationFixture());
  const render = spyOn(handle, 'renderPng').mockRejectedValue(Object.assign(new Error('canvas unavailable'), { code: 'canvas-unavailable' }));
  const layout = spyOn(handle, 'layoutSlide');
  try {
    const preview = await renderPptxSlide(handle, 1, 1.5);
    expect(render).toHaveBeenCalledTimes(1);
    expect(layout).toHaveBeenCalledWith(0);
    expect([...preview.png.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(preview.warnings).toContain('Shadow effects are omitted in server previews.');
  } finally { layout.mockRestore(); render.mockRestore(); handle.dispose(); }
});

test('PPTX previews propagate engine errors without using canvas', async () => {
  const handle = openPresentation(await presentationFixture());
  const render = spyOn(handle, 'renderPng');
  const layout = spyOn(handle, 'layoutSlide');
  try {
    for (const code of ['canvas-too-large', 'font-unavailable', 'engine-failed']) {
      const error = Object.assign(new Error('engine refused'), { code });
      render.mockRejectedValue(error);
      await expect(renderPptxSlide(handle, 1)).rejects.toBe(error);
    }
    expect(layout).not.toHaveBeenCalled();
  } finally { layout.mockRestore(); render.mockRestore(); handle.dispose(); }
});
