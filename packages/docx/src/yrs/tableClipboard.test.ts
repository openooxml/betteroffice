import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsCellLoc, type YrsSession } from './index';
import { storyOffsetForLoc } from './storyPlainText';
import {
  applyClipboardFragment,
  readClipboardSlice,
  validClipboardFragment,
  type ClipboardSliceTarget,
} from './clipboardFragment';

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

function cellAt(story: string, tableIndex: number, row: number, column: number): YrsCellLoc {
  return { story, tableIndex, row, column };
}

function fillCell(session: YrsSession, cell: YrsCellLoc, text: string, bold = false): string {
  const payload = session.tablePayload(cell.story, cell.tableIndex);
  const rows = (payload?.rows ?? []) as Array<{ cells: Array<{ story: string }> }>;
  const story = rows[cell.row]?.cells[cell.column]?.story;
  if (!story) throw new Error('cell story is missing');
  const paraId = session.paragraphs(story)[0]!.paraId;
  session.insertText({ story, paraId, offset: 0 }, text);
  if (bold) {
    session.formatRange(
      { story, start: { paraId, offset: 0 }, end: { paraId, offset: text.length } },
      { bold: true },
    );
  }
  return story;
}

function cellTexts(session: YrsSession, story: string, tableIndex: number): string[] {
  const payload = session.tablePayload(story, tableIndex);
  const rows = (payload?.rows ?? []) as Array<{ cells: Array<{ story: string }> }>;
  return rows.flatMap((row) =>
    row.cells.map((cell) =>
      session
        .storySegments(cell.story)
        .filter((segment) => segment.kind === 'text')
        .map((segment) => (segment.kind === 'text' ? segment.text : ''))
        .join(''),
    ),
  );
}

function tableCount(session: YrsSession, story: string): number {
  return session
    .storySegments(story)
    .filter((segment) => segment.kind === 'embed' && segment.embedKind === 'table').length;
}

async function tableSession(cells: string[][], boldFirst = false): Promise<{ session: YrsSession; table: { story: string; tableIndex: number } }> {
  const session = await createYrsSession({ clientId: 52001 });
  session.seedFromDocx(blankDocx());
  const anchor = session.paragraphs('body')[0]!;
  const receipt = session.insertTable({ story: 'body', paraId: anchor.paraId, offset: 0 }, cells.length, cells[0]!.length);
  cells.forEach((row, rowIndex) =>
    row.forEach((text, columnIndex) =>
      fillCell(session, cellAt('body', 0, rowIndex, columnIndex), text, boldFirst && rowIndex === 0 && columnIndex === 0),
    ),
  );
  return { session, table: receipt.table };
}

describe('table clipboard sessions', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('reads a whole table as text and a valid fragment', async () => {
    const { session } = await tableSession([['a', 'b'], ['c', 'd']], true);
    try {
      const target: ClipboardSliceTarget = {
        cells: {
          anchor: cellAt('body', 0, 0, 0),
          head: cellAt('body', 0, 1, 1),
        },
      };
      const slice = readClipboardSlice(session, target);
      expect(slice.text).toBe('a\tb\nc\td');
      expect(slice.fragment?.paragraphs).toHaveLength(1);
      expect(slice.fragment?.paragraphs[0]?.tables?.[0]?.rows).toHaveLength(2);
      expect(validClipboardFragment(slice.fragment)).toBe(true);
    } finally {
      session.destroy();
    }
  });

  it('expands a partial rect over a merged cell', async () => {
    const { session } = await tableSession([['a', 'b'], ['c', 'd']]);
    try {
      session.mergeCells({
        anchor: cellAt('body', 0, 0, 0),
        head: cellAt('body', 0, 0, 1),
      });
      const slice = readClipboardSlice(session, {
        cells: {
          anchor: cellAt('body', 0, 0, 1),
          head: cellAt('body', 0, 0, 1),
        },
      });
      expect(slice.fragment?.paragraphs[0]?.tables?.[0]?.rows).toHaveLength(1);
      expect(slice.fragment?.paragraphs[0]?.tables?.[0]?.rows[0]?.cells).toHaveLength(1);
      expect(slice.fragment?.paragraphs[0]?.tables?.[0]?.rows[0]?.cells[0]?.tcPr).toMatchObject({
        rowspan: 1,
        colspan: 2,
      });
      expect(validClipboardFragment(slice.fragment)).toBe(true);
    } finally {
      session.destroy();
    }
  });

  it('reads a body range spanning a table', async () => {
    const { session } = await tableSession([['a', 'b'], ['c', 'd']]);
    try {
      const first = session.paragraphs('body')[0]!;
      const slice = readClipboardSlice(session, {
        story: 'body',
        from: storyOffsetForLoc(session, { story: 'body', paraId: first.paraId, offset: 0 }),
        to: session.storyLength('body'),
      });
      expect(slice.text).toContain('a\tb\nc\td\n');
      expect(slice.fragment?.paragraphs.length).toBeGreaterThan(0);
      const tables = slice.fragment?.paragraphs.flatMap((paragraph) => paragraph.tables ?? []);
      expect(tables).toHaveLength(1);
      expect(tables?.[0]?.rows).toHaveLength(2);
      expect(validClipboardFragment(slice.fragment)).toBe(true);
    } finally {
      session.destroy();
    }
  });

  it('round trips copy and paste with identical content and one undo', async () => {
    const { session } = await tableSession([['a', 'b'], ['c', 'd']], true);
    try {
      const slice = readClipboardSlice(session, {
        cells: {
          anchor: cellAt('body', 0, 0, 0),
          head: cellAt('body', 0, 1, 1),
        },
      });
      expect(slice.fragment).not.toBeNull();
      const paragraphs = session.paragraphs('body');
      const last = paragraphs[paragraphs.length - 1]!;
      const end = session.locateParagraph('body', last.paraId).end;
      const start = session.locateParagraph('body', last.paraId).start;
      const caret = applyClipboardFragment(
        session,
        {
          story: 'body',
          start: { paraId: last.paraId, offset: end - start },
          end: { paraId: last.paraId, offset: end - start },
        },
        slice.fragment!,
      );
      expect(tableCount(session, 'body')).toBe(2);
      expect(cellTexts(session, 'body', 1)).toEqual(['a', 'b', 'c', 'd']);
      const first = session.tablePayload('body', 0);
      const second = session.tablePayload('body', 1);
      expect(second?.tblPr).toEqual(first?.tblPr);
      expect(second?.grid).toEqual(first?.grid);
      const pastedFirstCell = ((second?.rows ?? []) as Array<{ cells: Array<{ story: string }> }>)[0]?.cells[0]?.story;
      const attrs = session
        .storySegments(pastedFirstCell!)
        .find((segment) => segment.kind === 'text')!;
      expect(attrs.kind === 'text' && attrs.attributes.bold).toBe(true);
      expect(caret.story).toBe('body');
      expect(session.undo()).toBe(true);
      expect(tableCount(session, 'body')).toBe(1);
    } finally {
      session.destroy();
    }
  });

  it('round trips a 70-row table', async () => {
    const rows = Array.from({ length: 70 }, (_, row) => [`r${row}c0`, `r${row}c1`]);
    const { session } = await tableSession(rows);
    try {
      const slice = readClipboardSlice(session, {
        cells: {
          anchor: cellAt('body', 0, 0, 0),
          head: cellAt('body', 0, 69, 1),
        },
      });
      expect(slice.text.split('\n')).toHaveLength(70);
      const paragraphs = session.paragraphs('body');
      const last = paragraphs[paragraphs.length - 1]!;
      const end = session.locateParagraph('body', last.paraId).end;
      const start = session.locateParagraph('body', last.paraId).start;
      applyClipboardFragment(
        session,
        {
          story: 'body',
          start: { paraId: last.paraId, offset: end - start },
          end: { paraId: last.paraId, offset: end - start },
        },
        slice.fragment!,
      );
      expect(tableCount(session, 'body')).toBe(2);
      expect(cellTexts(session, 'body', 1)).toEqual(rows.flat());
    } finally {
      session.destroy();
    }
  });
});
