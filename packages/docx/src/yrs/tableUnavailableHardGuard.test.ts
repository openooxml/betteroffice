import { afterEach, beforeAll, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import type { YrsCellLoc, YrsSession } from './index';

const real = await import('./tableUnavailable');
mock.module('./tableUnavailable', () => ({
  ...real,
  hasUnavailableTable: () => false,
  refuseUnavailableTables: () => {},
}));
const gate = await import('./tableUnavailable');
const { createYrsSession, saveYrsDocx } = await import('./index');
const { yrsToDocument } = await import('./yrsToDocument');

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
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

beforeAll(async () => {
  preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  await preloadOpcWasm();
}, 60_000);

function exchange(a: YrsSession, b: YrsSession): void {
  const fromA = a.encodeState();
  const fromB = b.encodeState();
  b.applyUpdate(fromA);
  a.applyUpdate(fromB);
}

test('with the save gate switched off, converting an incomplete table still refuses with the code and writes nothing', async () => {
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
  expect(gate.hasUnavailableTable(a, base)).toBe(false);
  expect(() => gate.refuseUnavailableTables(a, base)).not.toThrow();

  let converted: unknown = 'not converted';
  try {
    converted = yrsToDocument(a, base);
  } catch (error) {
    expect((error as Error & { code?: string }).code).toBe('table-unavailable');
  }
  expect(converted).toBe('not converted');

  let written: unknown = 'not written';
  try {
    written = await saveYrsDocx(a, { updateModifiedDate: false });
  } catch (error) {
    expect((error as Error & { code?: string }).code).toBe('table-unavailable');
  }
  expect(written).toBe('not written');
});
