import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import type { LayoutBlock, TableBlock } from '../layout/pagination/types';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const margins = (sides: Record<string, number>) =>
  `<w:tblCellMar>${Object.entries(sides)
    .map(([side, twips]) => `<w:${side} w:w="${twips}" w:type="dxa"/>`)
    .join('')}</w:tblCellMar>`;

const table = (style: string, direct: Record<string, number>) =>
  `<w:tbl><w:tblPr><w:tblStyle w:val="${style}"/><w:tblW w:w="5000" w:type="pct"/>${margins(direct)}</w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`;

function source(): Uint8Array {
  const styles = `<w:styles ${W}><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>${margins({ top: 0, left: 108, bottom: 0, right: 108 })}</w:tblPr></w:style><w:style w:type="table" w:styleId="Narrow"><w:name w:val="Narrow"/><w:tblPr>${margins({ left: 54, right: 54 })}</w:tblPr></w:style></w:styles>`;
  const body = table('Missing', { top: 57, bottom: 57 }) + table('Narrow', { top: 57 });
  const parts: PartsMap = new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')],
    ['word/document.xml', toBytes(`<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`)],
    ['word/styles.xml', toBytes(styles)],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const round = (value: number) => Math.round(value * 1000) / 1000;
const px = (twips: number) => round(twips / 15);

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

test('table cell margins resolve each side through direct, table style and default table style', async () => {
  const bytes = source();
  const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  const native = await createYrsSession({ clientId: 75211 });
  const projected = await createYrsSession({ clientId: 75211 });
  try {
    native.seedFromDocx(bytes);
    documentToYrs(projected, parsed);
    expect(native.storySegments('body')).toEqual(projected.storySegments('body'));
    for (const session of [native, projected]) {
      const tables = (session.yrsBlocksForStory('body', {}) as LayoutBlock[]).filter(
        (block): block is TableBlock => block.kind === 'table'
      );
      const padding = tables.map((block) =>
        block.rows[0].cells.map((cell) =>
          Object.fromEntries(
            Object.entries(cell.padding ?? {}).map(([side, value]) => [side, round(value)])
          )
        )
      );
      expect(padding).toEqual([
        Array(2).fill({ top: px(57), right: px(108), bottom: px(57), left: px(108) }),
        Array(2).fill({ top: px(57), right: px(54), bottom: 0, left: px(54) }),
      ]);
    }
  } finally {
    native.destroy();
    projected.destroy();
  }
});
