import { beforeAll, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm, setEditWasmHeapLimit } from '../wasm/edit';
import { engineKind } from '../wasm/engineEdit';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentEngineWorkerClient, ResidentWorkerOutOfMemoryError } from './residentEngineWorkerClient';

const HEAP_LIMIT_BYTES = 64 * 1024;

let startWorker: Awaited<ReturnType<typeof residentWorkerFactory>>;

beforeAll(async () => {
  await preloadEditWasm(engineKind === 'native' ? undefined : new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
  )));
  startWorker = await residentWorkerFactory();
}, 30_000);

function docx(): Uint8Array<ArrayBuffer> {
  const paragraphs = Array.from({ length: 400 }, (_, index) =>
    `<w:p><w:r><w:t>Paragraph ${index} with enough text to hold some heap in the editing core.</w:t></w:r></w:p>`).join('');
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" ' +
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

test(`an open past the worker heap limit fails as out of memory on ${engineKind} and the client replaces the worker`, async () => {
  const bytes = docx();
  const worker = startWorker(74_200);
  const terminate = spyOn(worker, 'terminate');
  const client = new ResidentEngineWorkerClient(worker);
  try {
    const failure = await client.open(bytes, { heapLimitBytes: HEAP_LIMIT_BYTES }).then(
      () => undefined,
      (error: unknown) => error,
    );
    expect(failure).toBeInstanceOf(ResidentWorkerOutOfMemoryError);
    expect(client.hasFailed()).toBe(true);
    expect(terminate).toHaveBeenCalled();
  } finally {
    terminate.mockRestore();
    client.destroy();
    setEditWasmHeapLimit(undefined);
  }
  // A trapped wasm instance is shared by every in-process worker here; a real worker starts a new one.
  if (engineKind === 'wasm') return;
  const replacement = new ResidentEngineWorkerClient(startWorker(74_201));
  try {
    expect(typeof (await replacement.open(bytes)).hostJson).toBe('string');
  } finally {
    replacement.destroy();
  }
}, 30_000);
