import { expect, spyOn, test } from 'bun:test';
import JSZip from 'jszip';
import { isHostMessage } from '../../../../shared/office-session';
import type { WorkbookHandle } from '../wasm/loader';
import { hydratePeer, openWorkbookSession } from './client';

test('cold hydration instantiates exactly the worker module without main-thread wasm compilation', async () => {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>NOW()</f><v>0</v></c><c r="B1"><f>RANDBETWEEN(1,1000000)</f><v>0</v></c></row></sheetData></worksheet>');
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const compiling = spyOn(WebAssembly, 'compile');
  const compileStreaming = spyOn(WebAssembly, 'compileStreaming');
  const instantiateStreaming = spyOn(WebAssembly, 'instantiateStreaming');
  const instantiating = spyOn(WebAssembly, 'instantiate');
  const worker = new Worker(new URL('./worker.ts', import.meta.url), { type: 'module' });
  const modules: WebAssembly.Module[] = [];
  worker.addEventListener('message', (event: MessageEvent) => {
    if (isHostMessage(event.data) && event.data.kind === 'wasm-module') modules.push(event.data.module);
  });
  let session: Awaited<ReturnType<typeof openWorkbookSession>> | undefined;
  let peer: WorkbookHandle | undefined;
  try {
    session = await openWorkbookSession(bytes, { worker: () => worker, retainPeerHydration: true });
    expect(modules).toHaveLength(1);
    expect(modules[0]).toBeInstanceOf(WebAssembly.Module);
    expect(instantiating).not.toHaveBeenCalled();
    peer = await hydratePeer(session);
    expect(instantiating).toHaveBeenCalledTimes(1);
    expect(instantiating.mock.calls[0][0]).toBe(modules[0]);
    expect(instantiating.mock.calls.every(([input]) => input instanceof WebAssembly.Module)).toBe(true);
    expect(compiling).not.toHaveBeenCalled();
    expect(compileStreaming).not.toHaveBeenCalled();
    expect(instantiateStreaming).not.toHaveBeenCalled();
    expect(peer.version()).toBe(await session.call.version());
    expect(peer.save()).toEqual(await session.save());
    const cells = peer.readCells({
      ranges: [{ sheetId: 'sheet:0', range: { kind: 'a1', a1: 'A1:B1' } }],
    });
    if (!cells.ok) throw new Error(cells.failure.message);
    expect(cells.ranges[0].cells[0].map((cell) => cell.value.kind)).toEqual(['number', 'number']);
    for (const cell of cells.ranges[0].cells[0]) {
      if (cell.value.kind !== 'number') throw new Error('Missing volatile number');
      expect(cell.value.value).toBeGreaterThan(0);
    }
  } finally {
    peer?.dispose();
    await session?.dispose();
    worker.terminate();
    compiling.mockRestore();
    compileStreaming.mockRestore();
    instantiateStreaming.mockRestore();
    instantiating.mockRestore();
  }
});
