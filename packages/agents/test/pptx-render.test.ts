import { beforeAll, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { createFontProvider } from '@betteroffice/fonts';
import { initWasm, openPresentation, type SlideSnapshot } from '@betteroffice/pptx';
import { pptxFontFaces, pptxThemeFonts, renderPptxSlide } from '../src/pptx-render';
import { initialize } from '../src/prototype';
import { presentationFixture } from './pptx-fixture';

beforeAll(() => initialize('pptx', initWasm));

for (const failPrimary of [false, true]) {
  test(`PPTX preview skips rejected coverage with ${failPrimary ? 'a rejected' : 'a loaded'} primary face`, async () => {
    const zip = await JSZip.loadAsync(await presentationFixture());
    const path = 'ppt/slides/slide1.xml';
    zip.file(path, (await zip.file(path)!.async('string')).replaceAll('</a:rPr>', '<a:latin typeface="Segoe UI"/></a:rPr>'));
    const handle = openPresentation(await zip.generateAsync({ type: 'uint8array' }));
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
    } finally { handle.dispose(); }
  });
}

test('PPTX save faces follow the theme for runs that name no face', async () => {
  const zip = new JSZip();
  zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="s"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>');
  const theme = pptxThemeFonts(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  expect(theme).toEqual({ major: ['Calibri Light'], minor: ['Calibri'] });
  expect(pptxThemeFonts(await zip.generateAsync({ type: 'uint8array' }))).toEqual(theme);
  expect(pptxThemeFonts(new Uint8Array(64))).toEqual({ major: [], minor: [] });
  const prefixed = new JSZip();
  prefixed.file('ppt/theme/theme1.xml', '<d:theme xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><d:themeElements><d:fontScheme name="s"><d:majorFont><d:latin typeface="Georgia"/></d:majorFont><d:minorFont><d:latin typeface="Verdana"/></d:minorFont></d:fontScheme></d:themeElements></d:theme>');
  prefixed.file('ppt/theme/theme2.xml', '<theme xmlns="http://schemas.openxmlformats.org/drawingml/2006/main"><themeElements><fontScheme name="s"><majorFont><latin typeface="Tahoma"/></majorFont><minorFont><latin typeface="Verdana"/></minorFont></fontScheme></themeElements></theme>');
  expect(pptxThemeFonts(await prefixed.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))).toEqual({ major: ['Georgia', 'Tahoma'], minor: ['Verdana'] });
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
  expect(pptxThemeFonts(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))).toEqual({ major: ['Times New Roman'], minor: ['Arial'] });
});

test('PPTX theme fonts accept single-quoted typefaces', async () => {
  const zip = new JSZip();
  zip.file('ppt/theme/theme1.xml', `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme name="s"><a:majorFont><a:latin typeface='Times New Roman'/></a:majorFont><a:minorFont><a:latin typeface='Arial'/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`);
  expect(pptxThemeFonts(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))).toEqual({ major: ['Times New Roman'], minor: ['Arial'] });
});

test('PPTX previews measure runs without a face in the theme fonts the export uses', async () => {
  const handle = openPresentation(await presentationFixture());
  const registered: string[] = [];
  const tracked = { ...handle, registerFont: (face: Parameters<typeof handle.registerFont>[0]) => { registered.push(face.family); return handle.registerFont(face); } };
  try {
    const theme = { major: ['Georgia'], minor: ['Verdana'] };
    await renderPptxSlide(tracked, 1, 1, createFontProvider(), theme);
    const runs = handle.snapshot().slides[0].shapes.flatMap(shape => shape.textStories).flatMap(story => story.paragraphs).flatMap(paragraph => paragraph.runs);
    expect(runs.some(run => !run.style.fontFamily)).toBe(true);
    expect(registered).toEqual(expect.arrayContaining(['Georgia', 'Verdana']));
  } finally { handle.dispose(); }
});
