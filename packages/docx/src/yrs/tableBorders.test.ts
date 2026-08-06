import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import type { Document, Table } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import {
  createYrsSession,
  type YrsCellBorders,
  type YrsSession,
  type YrsTableRange,
} from './index';
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

const BLACK = { style: 'single', size: 4, color: { rgb: '000000' } };
const RED = { style: 'single', size: 4, color: { rgb: 'FF0000' } };
const CLEARED = { style: 'none', size: 0, color: { rgb: '000000' } };

/** The toolbar's own payloads. */
const ALL_BORDERS: YrsCellBorders = {
  top: BLACK,
  bottom: BLACK,
  left: BLACK,
  right: BLACK,
  insideH: BLACK,
  insideV: BLACK,
};
const INSIDE_BORDERS: YrsCellBorders = { insideH: BLACK, insideV: BLACK };

const SIDES = ['top', 'bottom', 'left', 'right'] as const;

function range(row: number, column: number, toRow = row, toColumn = column): YrsTableRange {
  return {
    anchor: { story: 'body', tableIndex: 0, row, column },
    head: { story: 'body', tableIndex: 0, row: toRow, column: toColumn },
  };
}

interface RenderedTable {
  kind: string;
  rows: Array<{
    cells: Array<{ borders?: Partial<Record<string, { style?: string; color?: string }>> }>;
  }>;
}

function renderedTable(session: YrsSession): RenderedTable {
  const blocks = session.yrsBlocksForStory('body') as RenderedTable[];
  const table = blocks.find((block) => block.kind === 'table');
  if (!table) throw new Error('no table in the rendered body');
  return table;
}

/** The edges `lower_cell_borders` actually hands the renderer, with colours. */
function paintedEdges(session: YrsSession, row: number, column: number): Record<string, string> {
  const borders = renderedTable(session).rows[row].cells[column].borders ?? {};
  const painted: Record<string, string> = {};
  for (const side of SIDES) {
    const edge = borders[side];
    if (edge && edge.style !== 'none') painted[side] = edge.color ?? '';
  }
  return painted;
}

function tableOf(document: Document): Table {
  const table = document.package.document.content.find((block) => block.type === 'table');
  if (!table) throw new Error('no table in document body');
  return table as Table;
}

async function griddedTable(clientId: number): Promise<{
  session: YrsSession;
  parsed: Document;
}> {
  const bytes = blankDocx();
  const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  const session = await createYrsSession({ clientId });
  session.seedFromDocx(bytes);
  const anchor = session.paragraphs('body')[0];
  session.insertTable({ story: 'body', paraId: anchor.paraId, offset: 0 }, 2, 2);
  return { session, parsed };
}

describe('the table border toolbar', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('leaves a cell’s other three sides alone when one side button is pressed', async () => {
    const { session } = await griddedTable(52001);
    try {
      session.setCellBorders(range(0, 0, 1, 1), ALL_BORDERS);
      expect(Object.keys(paintedEdges(session, 0, 0)).sort()).toEqual([
        'bottom',
        'left',
        'right',
        'top',
      ]);

      session.setCellBorders(range(0, 0), { top: RED });
      expect(paintedEdges(session, 0, 0)).toEqual({
        top: '#FF0000',
        bottom: '#000000',
        left: '#000000',
        right: '#000000',
      });
    } finally {
      session.destroy();
    }
  });

  it('clears one side without disturbing the rest', async () => {
    const { session } = await griddedTable(52002);
    try {
      session.setCellBorders(range(0, 0, 1, 1), ALL_BORDERS);
      session.setCellBorders(range(0, 0), { top: CLEARED });
      expect(Object.keys(paintedEdges(session, 0, 0)).sort()).toEqual(['bottom', 'left', 'right']);

      session.setCellBorders(range(0, 1), { top: null });
      expect(Object.keys(paintedEdges(session, 0, 1)).sort()).toEqual(['bottom', 'left', 'right']);
    } finally {
      session.destroy();
    }
  });

  it('renders inside borders as interior rules and saves no exterior ones', async () => {
    const { session, parsed } = await griddedTable(52003);
    let saved: Document;
    try {
      session.setCellBorders(range(0, 0, 1, 1), INSIDE_BORDERS);
      expect(Object.keys(paintedEdges(session, 0, 0)).sort()).toEqual(['bottom', 'right']);
      expect(Object.keys(paintedEdges(session, 1, 1)).sort()).toEqual(['left', 'top']);
      saved = yrsToDocument(session, parsed);
    } finally {
      session.destroy();
    }

    const borders = tableOf(saved).formatting?.borders;
    expect(borders?.insideH).toMatchObject(BLACK);
    expect(borders?.insideV).toMatchObject(BLACK);
    for (const side of SIDES) {
      expect(borders?.[side], `tblBorders.${side} must not be fabricated`).toBeUndefined();
    }
  });

  it('round-trips all borders through a save and reopen', async () => {
    const { session, parsed } = await griddedTable(52004);
    let saved: Document;
    try {
      session.setCellBorders(range(0, 0, 1, 1), ALL_BORDERS);
      saved = yrsToDocument(session, parsed);
    } finally {
      session.destroy();
    }

    const reopened = await parseDocx(await repackDocx(saved), { preloadFonts: false });
    const borders = tableOf(reopened).formatting?.borders;
    for (const side of [...SIDES, 'insideH', 'insideV'] as const) {
      expect(borders?.[side], `reopened tblBorders.${side}`).toMatchObject(BLACK);
    }
  });
});
