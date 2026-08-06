import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Document, Table } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';
import { yrsToDocument } from './yrsToDocument';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="${OFFICE_DOC}.wordprocessingml.styles+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Anchor</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;

/** No table style at all — the blank document the bug report starts from. */
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;

function blankDocx(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set('word/document.xml', toBytes(DOCUMENT));
  parts.set('word/styles.xml', toBytes(STYLES));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function onlyTable(document: Document): Table {
  const table = document.package.document.content.find((block) => block.type === 'table');
  if (!table) throw new Error('no table in document body');
  return table as Table;
}

const GRID_EDGE = { style: 'single', size: 4, color: { rgb: '000000' } };

describe('a table inserted into a document without a bordered table style', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('saves and reopens with Word default grid borders', async () => {
    const bytes = blankDocx();
    const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
    const session = await createYrsSession({ clientId: 51001 });

    let saved: Document;
    try {
      session.seedFromDocx(bytes);
      const anchor = session.paragraphs('body')[0];
      session.insertTable({ story: 'body', paraId: anchor.paraId, offset: 0 }, 2, 2);
      saved = yrsToDocument(session, parsed);
    } finally {
      session.destroy();
    }

    const table = onlyTable(saved);
    for (const side of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as const) {
      expect(table.formatting?.borders?.[side], `tblBorders.${side}`).toMatchObject(GRID_EDGE);
    }

    const reopenedBytes = await repackDocx(saved);
    const reopened = await parseDocx(reopenedBytes, { preloadFonts: false });
    const reopenedTable = onlyTable(reopened);
    for (const side of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as const) {
      expect(reopenedTable.formatting?.borders?.[side], `reopened tblBorders.${side}`).toMatchObject(
        GRID_EDGE
      );
    }

    const reseeded = await createYrsSession({ clientId: 51002 });
    try {
      reseeded.seedFromDocx(new Uint8Array(reopenedBytes));
      const blocks = reseeded.yrsBlocksForStory('body') as Array<{ kind: string }>;
      const block = blocks.find((entry) => entry.kind === 'table');
      expect(block, 'reseeded body still carries a table').toBeDefined();
      const { rows } = block as unknown as {
        rows: Array<{ cells: Array<{ borders?: Record<string, unknown> }> }>;
      };
      expect(rows).toHaveLength(2);
      for (const row of rows) {
        for (const cell of row.cells) {
          for (const side of ['top', 'bottom', 'left', 'right'] as const) {
            expect(cell.borders?.[side], `reseeded rendered ${side} rule`).toEqual({
              width: 1,
              color: '#000000',
              style: 'solid',
            });
          }
        }
      }
    } finally {
      reseeded.destroy();
    }
  });
});
