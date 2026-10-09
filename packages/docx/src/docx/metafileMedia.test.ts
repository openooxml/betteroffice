import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { LayoutBlock } from '../layout/pagination/types';
import { parseDocx } from '.';
import { repackDocx } from './rezip';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { readDocxContainer } from './zipContainer';

let wasm = true;
try {
  await import('./rustParseFacade');
} catch {
  wasm = false;
}
const describeIfWasm = wasm ? describe : describe.skip;

const EMF = new Uint8Array(
  readFileSync(resolve(import.meta.dir, '../../../../crates/ooxml-metafile/tests/fixtures/shapes.emf'))
);
const PART = 'word/media/image1.emf';

function brokenEmf(): Uint8Array {
  const bytes = EMF.slice();
  new DataView(bytes.buffer).setUint32(88, 250, true);
  return bytes;
}

function packageWith(emf: Uint8Array): ArrayBuffer {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Default Extension="emf" ContentType="image/x-emf"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>'
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>'
    )
  );
  parts.set(
    'word/_rels/document.xml.rels',
    toBytes(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>' +
        '</Relationships>'
    )
  );
  parts.set(PART, emf);
  parts.set(
    'word/document.xml',
    toBytes(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
        ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<w:body><w:p><w:r><w:drawing><wp:inline>' +
        '<wp:extent cx="4572000" cy="3048000"/><wp:docPr id="1" name="Picture 1"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic>' +
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'
    )
  );
  return rezipPartsToArrayBuffer(parts);
}

function findImages(value: unknown, found: Record<string, unknown>[] = []): Record<string, unknown>[] {
  if (Array.isArray(value)) {
    for (const item of value) findImages(item, found);
  } else if (value && typeof value === 'object') {
    const record = value as Record<string, unknown>;
    if (record.type === 'image') found.push(record);
    for (const child of Object.values(record)) findImages(child, found);
  }
  return found;
}

function svgOf(src: string): string {
  const prefix = 'data:image/svg+xml;base64,';
  expect(src.startsWith(prefix)).toBe(true);
  return Buffer.from(src.slice(prefix.length), 'base64').toString('utf8');
}

describeIfWasm('EMF media end to end', () => {
  test('renders an EMF picture as SVG while the saved package keeps the EMF bytes', async () => {
    const parsed = await parseDocx(packageWith(EMF), { preloadFonts: false });
    expect(parsed.warnings ?? []).toEqual([]);
    const images = findImages(parsed.package.document.content);
    expect(images).toHaveLength(1);
    expect(images[0]?.mimeType).toBe('image/svg+xml');
    const svg = svgOf(String(images[0]?.src));
    expect(svg.startsWith('<svg xmlns="http://www.w3.org/2000/svg"')).toBe(true);
    expect(svg).toContain('<path');

    const saved = readDocxContainer(await repackDocx(parsed));
    expect(Array.from(saved.file(PART) ?? [])).toEqual(Array.from(EMF));
    expect(saved.paths().some((path) => path.endsWith('.svg'))).toBe(false);
  });

  test('shows a placeholder and warns for an EMF the replay refuses', async () => {
    const parsed = await parseDocx(packageWith(brokenEmf()), { preloadFonts: false });
    expect(parsed.warnings).toEqual([
      expect.stringContaining(`EMF image ${PART} could not be converted for display`),
    ]);
    const images = findImages(parsed.package.document.content);
    expect(svgOf(String(images[0]?.src))).toContain('fill="#f1f3f4"');
    const saved = readDocxContainer(await repackDocx(parsed));
    expect(Array.from(saved.file(PART) ?? [])).toEqual(Array.from(brokenEmf()));
  });

  // The renderer paints the yrs-seeded src, which docx-edit parses itself.
  test('seeds the yrs image src the renderer paints as SVG', async () => {
    const { createYrsSession } = await import('../yrs');
    const session = await createYrsSession({ clientId: 75113 });
    try {
      session.seedFromDocx(new Uint8Array(packageWith(EMF)));
      const blocks = session.yrsBlocksForStory('body', {}) as LayoutBlock[];
      const runs = blocks.flatMap((block) => (block.kind === 'paragraph' ? block.runs : []));
      const image = runs.find((run) => run.kind === 'image');
      if (image?.kind !== 'image') throw new Error('missing image');
      expect(svgOf(image.src)).toContain('<path');
    } finally {
      session.destroy();
    }
  });
});
