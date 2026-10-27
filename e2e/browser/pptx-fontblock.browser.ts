import { test, expect, type Locator } from 'playwright/test';
import JSZip from 'jszip';

const fontFallbackMessage = /font faces failed to load; text uses fallback fonts\./i;
const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const types = (body: string) => `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${body}</Types>`;
const override = (path: string, type: string) => `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/>`;
const indices = [1, 2, 3];

async function presentation() {
  const zip = new JSZip();
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  zip.file('[Content_Types].xml', types(override('ppt/presentation.xml', 'presentationml.presentation.main') + indices.map(i => override(`ppt/slides/slide${i}.xml`, 'presentationml.slide')).join('')));
  zip.file('_rels/.rels', rels(`<Relationship Id="doc" Type="${r}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldIdLst>${indices.map(i => `<p:sldId id="${255 + i}" r:id="slide${i}"/>`).join('')}</p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', rels(indices.map(i => `<Relationship Id="slide${i}" Type="${r}/slide" Target="slides/slide${i}.xml"/>`).join('')));
  for (const i of indices) zip.file(`ppt/slides/slide${i}.xml`, `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld name="Slide ${i}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:solidFill><a:srgbClr val="172554"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Slide ${i} Arial text</a:t></a:r><a:endParaRPr sz="2400"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function painted(canvas: Locator) {
  return canvas.evaluate((element: HTMLCanvasElement) => {
    if (!element.width || !element.height) return '';
    const pixels = element.getContext('2d')!.getImageData(0, 0, element.width, element.height).data;
    for (let i = 0; i < pixels.length; i += 4) {
      if (pixels[i + 3]! > 0 && Math.min(pixels[i]!, pixels[i + 1]!, pixels[i + 2]!) < 200) return element.toDataURL();
    }
    return '';
  });
}

test.describe('pptx font downloads', () => {
  for (const blocked of [false, true]) {
    test(blocked ? 'opens and paints every slide with fonts blocked' : 'opens and paints every slide with fonts allowed', async ({ page }) => {
      test.setTimeout(60_000);
      const warnings: string[] = [];
      page.on('console', (message) => { if (message.type() === 'warning') warnings.push(message.text()); });
      const errors: Error[] = [];
      page.on('pageerror', (error) => errors.push(error));
      let aborted = 0;
      if (blocked) {
        await page.route('https://cdn.jsdelivr.net/**', (route) => { aborted++; return route.abort(); });
        await page.route(/\.(?:ttf|otf|woff2?)(?:\?.*)?$/, (route) => { aborted++; return route.abort(); });
      }
      await page.goto('/?format=pptx');
      await page.locator('input[type=file]').last().setInputFiles({
        name: 'fontblock.pptx', mimeType: 'application/octet-stream', buffer: await presentation(),
      });
      const entries = page.getByRole('complementary', { name: 'Slides', exact: true }).getByRole('button');
      const canvas = page.getByTestId('pptx-slide-canvas');
      await expect(entries).toHaveCount(indices.length, { timeout: 30_000 });
      const count = await entries.count();
      expect(count).toBe(indices.length);
      if (blocked) expect(aborted).toBeGreaterThan(0);
      let previous = '';
      for (let index = 0; index < count; index++) {
        await entries.nth(index).click();
        await expect.poll(async () => {
          const image = await painted(canvas);
          return image !== '' && image !== previous;
        }, { timeout: 5_000 }).toBe(true);
        previous = await painted(canvas);
      }
      expect(warnings.filter((message) => fontFallbackMessage.test(message))).toHaveLength(blocked ? 1 : 0);
      expect(errors).toEqual([]);
    });
  }
});
