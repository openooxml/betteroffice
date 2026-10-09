import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
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
