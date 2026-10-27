import { beforeAll, expect, test } from 'bun:test';
import JSZip from 'jszip';
import { createFontProvider } from '@betteroffice/fonts';
import { initWasm, openPresentation } from '@betteroffice/pptx';
import { renderPptxSlide } from '../src/pptx-render';
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
