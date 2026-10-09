import { expect, test } from 'bun:test';
import { createCanvas, loadImage } from '@napi-rs/canvas';
import JSZip from 'jszip';
import { openDocx } from '../src';
import { renderDocxPage } from '../src/render';
import { fixture } from './fixture';

async function imageFixture(data: Uint8Array, extension: string, mimeType: string) {
  const zip = await JSZip.loadAsync(await fixture());
  zip.file(`word/media/image.${extension}`, data);
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', `<Default Extension="${extension}" ContentType="${mimeType}"/></Types>`));
  zip.file('word/_rels/document.xml.rels', (await zip.file('word/_rels/document.xml.rels')!.async('string')).replace('</Relationships>', `<Relationship Id="picture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.${extension}"/></Relationships>`));
  const drawing = '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="picture"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  zip.file('word/document.xml', (await zip.file('word/document.xml')!.async('string')).replace('<w:body>', `<w:body>${drawing}`));
  return zip.generateAsync({ type: 'uint8array' });
}

test('renders actual before/after pages without changing the live document', async () => {
  const doc = await openDocx(await fixture(), { renderer: renderDocxPage });
  try {
    const [hit] = doc.grep({ query: 'Executive summary' }).matches;
    const proposal = doc.propose({ author: 'test', edits: [{ ...hit, oldText: hit.text, newText: 'Annual report 2026' }] });
    const before = await doc.render(1);
    const after = await doc.render(1, proposal.id);
    expect([...before.png.slice(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    expect(before.pageCount).toBeGreaterThan(0);
    expect(before.png).not.toEqual(after.png);
    expect(doc.read(hit.ref).text).toBe('Executive summary');
    await expect(doc.render(1000)).rejects.toThrow('Choose a page');
  } finally { doc.close(); }
});

test('renders embedded image pixels in live and proposed previews', async () => {
  const image = createCanvas(32, 32);
  const context = image.getContext('2d');
  context.fillStyle = '#ff0000';
  context.fillRect(0, 0, 32, 32);
  const doc = await openDocx(await imageFixture(await image.encode('png'), 'png', 'image/png'), { renderer: renderDocxPage });
  try {
    const [hit] = doc.grep({ query: 'Executive summary' }).matches;
    const proposal = doc.propose({ author: 'test', edits: [{ match: hit.match, newText: 'Annual report' }] });
    for (const id of [undefined, proposal.id]) {
      const rendered = await doc.render(1, id);
      expect(rendered.warnings.filter(warning => warning.startsWith('Image'))).toEqual([]);
      const preview = createCanvas(rendered.width, rendered.height);
      const pixels = preview.getContext('2d');
      pixels.drawImage(await loadImage(Buffer.from(rendered.png)), 0, 0);
      const data = pixels.getImageData(0, 0, preview.width, preview.height).data;
      let redPixels = 0;
      for (let i = 0; i < data.length; i += 4) if (data[i] > 240 && data[i + 1] < 10 && data[i + 2] < 10) redPixels++;
      expect(redPixels).toBeGreaterThan(8000);
    }
  } finally { doc.close(); }
});

test('skips oversized rasters and SVG before native image decoding', async () => {
  const png = await createCanvas(32, 32).encode('png');
  png.writeUInt32BE(30000, 16);
  png.writeUInt32BE(30000, 20);
  const svg = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32"><image href="file:///tmp/outside.png" width="32" height="32"/></svg>');
  for (const [bytes, extension, type, warning] of [
    [png, 'png', 'image/png', 'An embedded image exceeds the preview image limits.'],
    [svg, 'svg', 'image/svg+xml', 'An embedded image uses an unsupported encoding; previews support raster images only.'],
  ] as const) {
    const doc = await openDocx(await imageFixture(bytes, extension, type), { renderer: renderDocxPage });
    try {
      expect((await doc.render(1)).warnings).toContain(warning);
    } finally { doc.close(); }
  }
});
