import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import type { SessionTransport } from '../../../../shared/office-session';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { isClientMessage } from '../../../../shared/office-session/protocol';
import { initWasm } from '../wasm/loader';
import { createWorkbookSession, type OpenWorkbookSessionOptions, type WorkbookSession } from './client';
import { createWorkbookSessionHost } from './host';
import type { WorkbookCommand, WorkbookSharedCommand } from './commandProtocol';

export function capturedWorkbookCommand(message: unknown): WorkbookCommand | undefined {
  if (!isClientMessage(message) || message.kind !== 'call') return undefined;
  if (message.method === 'command') return message.args[0] as WorkbookCommand;
  if (message.method === 'edit') return (message.args[0] as WorkbookSharedCommand).payload.compatibility;
  return undefined;
}

export async function loadWorkbookSessionFixtures(): Promise<{
  fixture: Uint8Array;
  chartFixture: Uint8Array;
  wasmBytes: Uint8Array<ArrayBuffer>;
}> {
  const [wasm, xlsx, charts] = await Promise.all([
    readFile(resolve(import.meta.dir, '../wasm/generated/xlsx_wasm_bg.wasm')),
    readFile(resolve(import.meta.dir, '../../test-fixtures/sample.xlsx')),
    readFile(resolve(import.meta.dir, '../../test-fixtures/charts.xlsx')),
  ]);
  const wasmBytes = new Uint8Array(wasm);
  await initWasm(wasmBytes);
  return { wasmBytes, fixture: new Uint8Array(xlsx), chartFixture: new Uint8Array(charts) };
}

export async function createTestWorkbookSession(
  bytes: Uint8Array,
  wrapHost?: (transport: SessionTransport) => SessionTransport,
  options: OpenWorkbookSessionOptions = {}
): Promise<WorkbookSession> {
  const pair = createInProcessPair();
  createWorkbookSessionHost(wrapHost ? wrapHost(pair.host) : pair.host);
  return createWorkbookSession(bytes, options, pair.client);
}

export function batchRequests(transport: SessionTransport): SessionTransport {
  return { ...transport, listen(listener) {
    let messages: unknown[] = [];
    let timer: ReturnType<typeof setTimeout> | undefined;
    const off = transport.listen((message) => {
      messages.push(message);
      if (timer !== undefined) return;
      timer = setTimeout(() => {
        timer = undefined;
        const batch = messages;
        messages = [];
        for (const queued of batch) listener(queued);
      }, 0);
    });
    return () => {
      off();
      if (timer !== undefined) clearTimeout(timer);
      messages = [];
    };
  } };
}

const memories: WebAssembly.Memory[] = [];

/** One sheet of number rows, large enough that opening it allocates past the engine's free heap. */
export async function numbersWorkbook(rows: number): Promise<Uint8Array> {
  const data = Array.from({ length: rows }, (_, row) =>
    `<row r="${row + 1}"><c r="A${row + 1}"><v>${row}</v></c><c r="B${row + 1}"><v>${row * 2}</v></c></row>`);
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${data.join('')}</sheetData></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

/** Grows the engine memory to the wasm32 ceiling, so its next allocation fails the way a huge workbook's does. */
export function exhaustEngineMemory(): void {
  if (memories.length === 0) throw new Error('Engine memory was not recorded; call initRecordedWasm first');
  for (const memory of memories) {
    for (let pages = 1 << 15; pages >= 1; pages >>= 1) {
      for (;;) {
        try { memory.grow(pages); } catch { break; }
      }
    }
  }
}

export async function initRecordedWasm(): Promise<void> {
  const instantiate = WebAssembly.instantiate;
  WebAssembly.instantiate = (async (...args: Parameters<typeof WebAssembly.instantiate>) => {
    const result = await (instantiate as (...input: unknown[]) => Promise<unknown>).apply(WebAssembly, args);
    const instance = result instanceof WebAssembly.Instance ? result : (result as { instance: WebAssembly.Instance }).instance;
    for (const value of Object.values(instance.exports)) if (value instanceof WebAssembly.Memory) memories.push(value);
    return result;
  }) as typeof WebAssembly.instantiate;
  try {
    await initWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/xlsx_wasm_bg.wasm'))));
  } finally {
    WebAssembly.instantiate = instantiate;
  }
}
