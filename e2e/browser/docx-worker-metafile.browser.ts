import { expect, test, type Page } from 'playwright/test';
import JSZip from 'jszip';

const le = (bits: 16 | 32, values: number[]) => {
  const buffer = Buffer.alloc(values.length * (bits / 8));
  values.forEach((value, index) => bits === 16 ? buffer.writeInt16LE(value, index * 2) : buffer.writeInt32LE(value, index * 4));
  return buffer;
};
const colorref = ([r, g, b]: number[]) => r! | (g! << 8) | (b! << 16);

/** A 100 by 50 pixel EMF holding one rectangle filled with `color`. */
function emf(color: number[]) {
  const records = [
    [39, le(32, [1, 0, colorref(color), 0])],
    [37, le(32, [1])],
    [43, le(32, [0, 0, 100, 50])],
  ] as const;
  const body = Buffer.concat(records.map(([kind, data]) => Buffer.concat([le(32, [kind, data.length + 8]), data])));
  const header = Buffer.concat([
    le(32, [1, 88, 0, 0, 99, 49, 0, 0, 2646, 1323, 0x464d4520, 0x10000, 0, records.length + 2]),
    le(16, [16, 0]), le(32, [0, 0, 0, 9600, 9600, 2540, 2540]),
  ]);
  const out = Buffer.concat([header, body, le(32, [14, 20, 0, 16, 20])]);
  out.writeUInt32LE(out.length, 48);
  return out;
}

/** A one by half inch placeable WMF holding one rectangle filled with `color`. */
function wmf(color: number[]) {
  const records = [
    [0x02fc, Buffer.concat([le(16, [0]), le(32, [colorref(color)]), le(16, [0])])],
    [0x012d, le(16, [0])],
    [0x041b, le(16, [720, 1440, 0, 0])],
  ] as const;
  const body = Buffer.concat(records.map(([fn, params]) => Buffer.concat([le(32, [params.length / 2 + 3]), le(16, [fn]), params])));
  const placeable = Buffer.concat([le(32, [0x9ac6cdd7 | 0]), le(16, [0, 0, 0, 1440, 720, 1440]), le(32, [0])]);
  let checksum = 0;
  for (let index = 0; index < placeable.length; index += 2) checksum ^= placeable.readUInt16LE(index);
  const total = (18 + body.length + 6) / 2;
  return Buffer.concat([
    placeable, le(16, [checksum]), le(16, [1, 9, 0x0300]), le(32, [total]), le(16, [16]), le(32, [8]), le(16, [0]),
    body, le(32, [3]), le(16, [0]),
  ]);
}

const svg = ([r, g, b]: number[]) => Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="96" height="48" viewBox="0 0 96 48"><rect width="96" height="48" fill="rgb(${r},${g},${b})"/></svg>`);

const pictures = [
  { name: 'EMF', target: 'media/image1.emf', color: [208, 32, 48], bytes: emf },
  { name: 'WMF', target: 'media/image2.wmf', color: [32, 64, 208], bytes: wmf },
  { name: 'SVG', target: 'media/image3.svg', color: [32, 160, 64], bytes: svg },
];

async function open(page: Page) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="emf" ContentType="image/x-emf"/><Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="svg" ContentType="image/svg+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const paragraph = (index: number) => `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="${index + 1}" name="Picture ${index + 1}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${index + 1}" name="Picture ${index + 1}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage${index}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${pictures.map((_, index) => paragraph(index)).join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${pictures.map((picture, index) => `<Relationship Id="rIdImage${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${picture.target}"/>`).join('')}</Relationships>`);
  for (const picture of pictures) zip.file(`word/${picture.target}`, picture.bytes(picture.color));
  await page.goto('/?format=docx');
  await page.locator('input[type=file]').last().setInputFiles({ name: 'metafiles.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await zip.generateAsync({ type: 'nodebuffer' }) });
  const canvas = page.locator('.paged-editor__page canvas').first();
  await expect(canvas).toBeVisible({ timeout: 60000 });
  return canvas;
}

/** Pixels of the page canvas within a small distance of each picture's fill colour. */
async function inkByColor(page: Page, png: Buffer) {
  return page.evaluate(async ({ data, colors }) => {
    const image = new Image();
    image.src = `data:image/png;base64,${data}`;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = image.naturalWidth;
    canvas.height = image.naturalHeight;
    const context = canvas.getContext('2d')!;
    context.drawImage(image, 0, 0);
    const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
    return colors.map((color) => {
      let count = 0;
      for (let index = 0; index < pixels.length; index += 4) {
        if (color.every((channel, offset) => Math.abs(pixels[index + offset]! - channel) <= 24)) count += 1;
      }
      return count;
    });
  }, { data: png.toString('base64'), colors: pictures.map(picture => picture.color) });
}

test('EMF, WMF and SVG pictures paint in the default editor', async ({ page }) => {
  const canvas = await open(page);
  await expect.poll(async () => {
    const ink = await inkByColor(page, await canvas.screenshot());
    return Object.fromEntries(pictures.map((picture, index) => [picture.name, ink[index]! >= 500]));
  }, { timeout: 30000 }).toEqual({ EMF: true, WMF: true, SVG: true });
});
