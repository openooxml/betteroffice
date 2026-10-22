import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { createYrsSession, saveYrsDocx, type YrsSession, type YrsTableRange } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
const table = (...cells: string[]) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${cells
    .map(() => '<w:gridCol w:w="4680"/>')
    .join('')}</w:tblGrid><w:tr>${cells.map(cell).join('')}</w:tr></w:tbl>`;

function docx(body: string): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(CONTENT_TYPES));
  parts.set('_rels/.rels', toBytes(PACKAGE_RELS));
  parts.set(
    'word/document.xml',
    toBytes(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`
    )
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const FIRST_ROW: YrsTableRange = {
  anchor: { story: 'body', tableIndex: 0, row: 0, column: 0 },
  head: { story: 'body', tableIndex: 0, row: 0, column: 1 },
};

async function opened(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession();
  session.seedFromDocx(bytes);
  return session;
}

async function reopenedStoryTexts(bytes: Uint8Array): Promise<Record<string, string[]>> {
  const session = await opened(bytes);
  try {
    return Object.fromEntries(
      session
        .storyIds()
        .filter((story) => story !== 'body' && story.startsWith('body'))
        .map((story) => [story, session.paragraphs(story).map((p) => p.text)])
    );
  } finally {
    session.destroy();
  }
}

describe('saving merged table cells after an earlier save', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('keeps the merged-in cell text', async () => {
    const session = await opened(
      docx(paragraph('Before') + table(paragraph('CellA'), paragraph('CellB')) + paragraph('After'))
    );
    try {
      await saveYrsDocx(session);
      session.mergeCells(FIRST_ROW);
      const saved = await saveYrsDocx(session);
      expect(await reopenedStoryTexts(saved.bytes)).toEqual({ 'body:t0:r0c0': ['CellA', 'CellB'] });
    } finally {
      session.destroy();
    }
  });

  it('keeps edits to a nested table the merge moved into the surviving cell', async () => {
    const session = await opened(
      docx(table(paragraph('CellA'), table(paragraph('Inner')) + paragraph('')))
    );
    try {
      const nested = 'body:t0:r0c1:t0:r0c0';
      expect(session.paragraphs(nested).map((p) => p.text)).toEqual(['Inner']);
      await saveYrsDocx(session);
      session.mergeCells(FIRST_ROW);
      await saveYrsDocx(session);
      const [inner] = session.paragraphs(nested);
      session.insertText({ story: nested, paraId: inner!.paraId, offset: 5 }, ' edited');
      const saved = await saveYrsDocx(session);
      expect(await reopenedStoryTexts(saved.bytes)).toEqual({
        'body:t0:r0c0': ['CellA', ''],
        'body:t0:r0c0:t0:r0c0': ['Inner edited'],
      });
    } finally {
      session.destroy();
    }
  });
});

describe('saving deleted table cells', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('retains deleted cells through a save until a fresh seed replaces the session', async () => {
    const session = await opened(docx(table(paragraph('CellA')) + paragraph('After')));
    try {
      session.deleteTable({ story: 'body', tableIndex: 0 });
      const saved = await saveYrsDocx(session);
      expect(new TextDecoder().decode(session.encodeState())).toContain('CellA');
      expect(session.storyIds()).toEqual(['body']);
      const reseeded = await opened(saved.bytes);
      try {
        expect(new TextDecoder().decode(reseeded.encodeState())).not.toContain('CellA');
        expect(reseeded.storyIds()).toEqual(['body']);
      } finally {
        reseeded.destroy();
      }
    } finally {
      session.destroy();
    }
  });

  for (const removal of ['table', 'row', 'column'] as const) {
    it(`excludes deleted cells after deleting a ${removal}, undoing and redoing`, async () => {
      const secondRow = `<w:tr>${cell(paragraph('CellC'))}${cell(paragraph('CellD'))}</w:tr>`;
      const content = table(paragraph('CellA'), paragraph('CellB')).replace(
        '</w:tbl>',
        `${secondRow}</w:tbl>`
      );
      const session = await opened(docx(paragraph('Before') + content + paragraph('After')));
      try {
        await saveYrsDocx(session);
        const first = { story: 'body', tableIndex: 0, row: 0, column: 0 };
        if (removal === 'table') session.deleteTable(first);
        else if (removal === 'row') session.deleteRow({ anchor: first, head: first });
        else session.deleteColumn({ anchor: first, head: first });
        const expected = {
          table: [],
          row: [['CellC'], ['CellD']],
          column: [['CellB'], ['CellD']],
        }[removal];
        for (let cycle = 0; cycle < 2; cycle += 1) {
          const saved = await saveYrsDocx(session);
          expect(Object.values(await reopenedStoryTexts(saved.bytes))).toEqual(expected);
          expect(session.undo()).toBe(true);
          const restored = await saveYrsDocx(session);
          expect(Object.values(await reopenedStoryTexts(restored.bytes))).toEqual([
            ['CellA'],
            ['CellB'],
            ['CellC'],
            ['CellD'],
          ]);
          expect(session.redo()).toBe(true);
        }
      } finally {
        session.destroy();
      }
    });
  }
});

describe('saving table rows with grid offsets', () => {
  beforeAll(async () => {
    await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
    await preloadOpcWasm(new Uint8Array(readFileSync(
      resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm')
    )));
  });

  it('keeps row grid offsets and their widths when another row is rewritten', async () => {
    const offsets = '<w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>' +
      '<w:wBefore w:w="4680" w:type="dxa"/><w:wAfter w:w="4680" w:type="dxa"/></w:trPr>';
    const grid = '<w:gridCol w:w="4680"/>'.repeat(4);
    const body = `<w:tbl><w:tblPr/><w:tblGrid>${grid}</w:tblGrid>` +
      `<w:tr>${['Left', 'A', 'B', 'Right'].map((text) => cell(paragraph(text))).join('')}</w:tr>` +
      `<w:tr>${offsets}${cell(paragraph('C'))}${cell(paragraph('D'))}</w:tr></w:tbl>${paragraph('Tail')}`;
    const session = await opened(docx(body));
    try {
      session.mergeCells({
        anchor: { story: 'body', tableIndex: 0, row: 0, column: 1 },
        head: { story: 'body', tableIndex: 0, row: 0, column: 2 },
      });
      const saved = await saveYrsDocx(session, { updateModifiedDate: false });
      const xml = new TextDecoder().decode(unzipContainer(saved.bytes)['word/document.xml']);
      const rows = xml.match(/<w:tr\b[\s\S]*?<\/w:tr>/g)!;
      expect(rows).toHaveLength(2);
      expect(rows[0].match(/<w:tc>/g)).toHaveLength(3);
      expect(rows[0]).toContain('<w:gridSpan w:val="2"/>');
      expect(rows[1]).toContain(offsets);
      expect(rows[1].match(/<w:tc>/g)).toHaveLength(2);
      expect(rows[1]).toMatch(/>C<[\s\S]*>D</);
      expect(Object.values(await reopenedStoryTexts(saved.bytes)).flat().filter(Boolean).sort()).toEqual(
        ['A', 'B', 'C', 'D', 'Left', 'Right']
      );
    } finally {
      session.destroy();
    }
  });
});
