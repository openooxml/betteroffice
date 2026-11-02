import { expect, test, type Locator } from 'playwright/test';
import JSZip from 'jszip';

const latin = 'Packed Latin regular and bold';
const devanagari = 'नमस्ते';
const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const types = (body: string) => `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${body}</Types>`;
const override = (path: string, type: string) => `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/>`;

async function documentBytes() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', types(override('word/document.xml', 'wordprocessingml.document.main')));
  zip.file('_rels/.rels', rels(`<Relationship Id="doc" Type="${relationships}/officeDocument" Target="word/document.xml"/>`));
  const runs = ['Calibri', 'Mangal'].flatMap(family => [false, true].map(bold =>
    `<w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:cs="${family}"/>${bold ? '<w:b/><w:bCs/>' : ''}<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr><w:t xml:space="preserve">${family === 'Calibri' ? latin : devanagari} </w:t></w:r>`)).join('');
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${runs}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function presentationBytes() {
  const zip = new JSZip();
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  zip.file('[Content_Types].xml', types(override('ppt/presentation.xml', 'presentationml.presentation.main') + override('ppt/slides/slide1.xml', 'presentationml.slide')));
  zip.file('_rels/.rels', rels(`<Relationship Id="doc" Type="${relationships}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${relationships}"><p:sldIdLst><p:sldId id="256" r:id="slide1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', rels(`<Relationship Id="slide1" Type="${relationships}/slide" Target="slides/slide1.xml"/>`));
  const paragraphs = ['Calibri', 'Mangal'].flatMap(family => [false, true].map(bold =>
    `<a:p><a:r><a:rPr sz="2400" b="${bold ? 1 : 0}"><a:solidFill><a:srgbClr val="172554"/></a:solidFill><a:latin typeface="${family}"/><a:cs typeface="${family}"/></a:rPr><a:t>${family === 'Calibri' ? latin : devanagari}</a:t></a:r><a:endParaRPr sz="2400"/></a:p>`)).join('');
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="3657600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function painted(canvas: Locator): Promise<boolean> {
  return canvas.evaluate((element: HTMLCanvasElement) => {
    if (!element.width || !element.height) return false;
    const pixels = element.getContext('2d')!.getImageData(0, 0, element.width, element.height).data;
    for (let index = 0; index < pixels.length; index += 4) {
      if (pixels[index + 3]! > 0 && Math.min(pixels[index]!, pixels[index + 1]!, pixels[index + 2]!) < 128) return true;
    }
    return false;
  });
}

for (const format of ['docx', 'pptx'] as const) {
  test(`${format} paints with packed fonts and only its engine decoder`, async ({ page }) => {
    test.setTimeout(90_000);
    const requests: string[] = [];
    const errors: string[] = [];
    const warnings: string[] = [];
    let lastRequestStarted = Date.now();
    page.on('request', request => {
      lastRequestStarted = Date.now();
      requests.push(request.url());
    });
    page.on('pageerror', error => errors.push(error.message));
    page.on('console', message => {
      if (message.type() === 'warning' && /font.*fail|fontRegistry.*falling back/i.test(message.text())) warnings.push(message.text());
    });
    await page.goto(`/?format=${format}`);
    await page.locator('input[type=file]').last().setInputFiles({
      name: `single-decoder.${format}`, mimeType: 'application/octet-stream',
      buffer: format === 'docx' ? await documentBytes() : await presentationBytes(),
    });
    const canvas = format === 'docx' ? page.locator('.paged-editor__page canvas').first() : page.getByTestId('pptx-slide-canvas');
    await expect(canvas).toBeVisible({ timeout: 45_000 });
    if (format === 'docx') {
      await expect(page.locator('[data-docx-worker-editor]')).toHaveAttribute('aria-busy', 'false', { timeout: 45_000 });
      await expect(page.locator('.canvas-page-mirror').first()).toContainText(latin);
      await expect(page.locator('.canvas-page-mirror').first()).toContainText(devanagari);
    } else {
      await expect(page.getByRole('complementary', { name: 'Slides', exact: true }).getByRole('button'))
        .toHaveCount(1, { timeout: 45_000 });
    }
    await expect.poll(() => painted(canvas), { timeout: 30_000 }).toBe(true);
    const paintedAt = Date.now();
    await expect.poll(() => Date.now() - Math.max(paintedAt, lastRequestStarted), { timeout: 30_000 }).toBeGreaterThanOrEqual(3_000);
    expect(requests.some(url => new URL(url).pathname.endsWith('.zst'))).toBe(true);
    const faces = format === 'docx'
      ? ['Carlito-Regular', 'Carlito-Bold', 'NotoSansDevanagari-Regular'] : ['LiberationSans-Regular'];
    for (const face of faces) {
      const asset = new RegExp(`/${face}\\.ttf(-[\\w-]+)?\\.zst$`);
      expect(requests.some(url => asset.test(new URL(url).pathname)), face).toBe(true);
    }
    expect(requests.filter(url => /\/fontpack(-[\w-]+)?\.wasm$/.test(new URL(url).pathname))).toEqual([]);
    expect(errors).toEqual([]);
    expect(warnings).toEqual([]);
  });
}
