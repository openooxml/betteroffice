import { afterEach, beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '@betteroffice/docx/docx/rezip/parts';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, yrsToDocument, type YrsCellLoc, type YrsSession } from '@betteroffice/docx/yrs';
import { projectForSave } from './useYrsCoreSession';

const ROOT = resolve(import.meta.dir, '../../../../../..');
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;

function docx(): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/></Types>`));
  parts.set('_rels/.rels', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`));
  parts.set('word/document.xml', toBytes(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraph('Before')}<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid><w:tr>${cell('First')}${cell('Second')}</w:tr></w:tbl>${paragraph('After')}<w:sectPr/></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const owned: YrsSession[] = [];
afterEach(() => {
  for (const session of owned.splice(0)) session.destroy();
});

beforeAll(() =>
  preloadEditWasm(new Uint8Array(readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm')))),
30_000);

function exchange(a: YrsSession, b: YrsSession): void {
  const fromA = a.encodeState();
  const fromB = b.encodeState();
  b.applyUpdate(fromA);
  a.applyUpdate(fromB);
}

test('the editor keeps typing while a table is unavailable: each change projects once, refused quietly, and saves resume after the table is deleted', async () => {
  const source = docx();
  const a = await createYrsSession({ clientId: 61 });
  const b = await createYrsSession({ clientId: 62 });
  owned.push(a, b);
  a.seedFromDocx(source);
  b.openDocx(source, false);
  b.applyUpdate(a.encodeState());
  const at: YrsCellLoc = { story: 'body', tableIndex: 0, row: 0, column: 0 };
  a.insertRow(at, 'below');
  exchange(a, b);
  b.setColumnWidth(at, 2000);
  expect(a.undo()).toBe(true);
  exchange(a, b);
  expect(b.undo()).toBe(true);
  exchange(a, b);

  const base = a.materializeDocx()!;
  const reported = { current: false };
  let attempts = 0;
  const project = () => {
    attempts += 1;
    return yrsToDocument(a, base);
  };
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  const error = spyOn(console, 'error').mockImplementation(() => {});
  try {
    for (const typed of ['x', 'y', 'z']) {
      const [first] = a.paragraphs('body');
      a.insertText({ story: 'body', paraId: first!.paraId, offset: first!.text.length }, typed);
      expect(projectForSave(project, reported)).toBeNull();
    }
    expect(attempts).toBe(3);
    expect(a.paragraphs('body')[0]!.text).toBe('Beforexyz');
    expect(warn.mock.calls.filter(([message]) => String(message).includes('table-unavailable'))).toHaveLength(1);
    expect(error).not.toHaveBeenCalled();

    a.deleteTable({ story: 'body', tableIndex: 0 });
    expect(projectForSave(project, reported)).not.toBeNull();
    expect(reported.current).toBe(false);
    expect(attempts).toBe(4);
  } finally {
    warn.mockRestore();
    error.mockRestore();
  }
});
