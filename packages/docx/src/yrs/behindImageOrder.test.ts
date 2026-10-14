import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import fixture from '../../../../crates/docx-edit/tests/fixtures/behind_objects.json';
import { parseDocx } from '../docx';
import type { LayoutBlock } from '../layout/pagination/types';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

function source(rank: number, shapeType?: string): Uint8Array {
  const image = fixture.image.replace('IMAGE_RANK', String(rank)).replace('IMAGE_BEHIND', '1')
    .replace('prst="rect"', `prst="${shapeType ?? 'rect'}"`);
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/></Relationships>')],
    ['word/document.xml', toBytes(`<w:document ${fixture.namespaces}><w:body>${image}${fixture.body}</w:body></w:document>`)],
    ['word/media/image.png', new Uint8Array(fixture.png)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

test.each([
  { rank: 0, shapeType: undefined },
  { rank: 20, shapeType: 'ellipse' },
  { rank: 4_294_967_295, shapeType: 'roundRect' },
  { rank: -1, shapeType: 'ellipse' },
  { rank: 1.5, shapeType: undefined },
  { rank: 4_294_967_296, shapeType: undefined },
])('both import paths retain image geometry and drawing order %j', async ({ rank, shapeType }) => {
  const bytes = source(rank, shapeType);
  const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  const native = await createYrsSession({ clientId: 75111 });
  const projected = await createYrsSession({ clientId: 75111 });
  try {
    native.seedFromDocx(bytes);
    documentToYrs(projected, parsed);
    expect(native.storySegments('body')).toEqual(projected.storySegments('body'));
    for (const session of [native, projected]) {
      const blocks = session.yrsBlocksForStory('body', {}) as LayoutBlock[];
      const image = blocks.flatMap((block) => block.kind === 'paragraph' ? block.runs : [])
        .find((run) => run.kind === 'image');
      expect(image?.kind).toBe('image');
      if (image?.kind !== 'image') throw new Error('missing image');
      expect(image.position?.relativeHeight).toBe(
        Number.isInteger(rank) && rank >= 0 && rank <= 4_294_967_295 ? rank : undefined
      );
      expect(image.position?.horizontal?.posOffset).toBe(914400);
      expect(image.position?.vertical?.posOffset).toBe(914400);
      expect(image.shapeType).toBe(shapeType);
      const saved = yrsToDocument(session, parsed);
      const paragraph = saved.package.document.content[0];
      if (paragraph?.type !== 'paragraph') throw new Error('missing saved paragraph');
      const run = paragraph.content.find((content) => content.type === 'run');
      if (run?.type !== 'run') throw new Error('missing saved run');
      const savedImage = run.content.find((content) => content.type === 'drawing');
      if (savedImage?.type !== 'drawing' || !savedImage.image) throw new Error('missing saved image');
      expect(savedImage.image.position?.relativeHeight).toBe(rank);
      expect(savedImage.image.shapeType).toBe(shapeType);
    }
  } finally {
    native.destroy();
    projected.destroy();
  }
});
