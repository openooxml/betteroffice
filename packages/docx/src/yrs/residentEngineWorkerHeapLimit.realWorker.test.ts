import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { exportDocxMarkdown } from '../core';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { engineKind } from '../wasm/engineEdit';
import { DocxSourceWorkerClient } from '../worker';
import { createResidentEngineSession } from './residentEngineSession';

const HEAP_LIMIT_BYTES = 64 * 1024;
const ROOMY_HEAP_LIMIT_BYTES = 256 * 1024 * 1024;
const OUT_OF_MEMORY = /^Document worker ran out of memory allocating \d+ bytes/;
const LAST_PARAGRAPH = 'Paragraph 399 with enough text to hold some heap in the editing core.';

const cleanups: (() => void)[] = [];

beforeAll(async () => {
  await preloadEditWasm(engineKind === 'native' ? undefined : new Uint8Array(readFileSync(
    resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
  )));
}, 30_000);

afterEach(() => {
  const errors: unknown[] = [];
  for (const cleanup of cleanups.splice(0).reverse()) {
    try {
      cleanup();
    } catch (error) {
      errors.push(error);
    }
  }
  if (errors.length === 1) throw errors[0];
  if (errors.length > 1) throw new AggregateError(errors, 'heap limit test cleanups failed');
});

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

function client(): DocxSourceWorkerClient {
  const created = new DocxSourceWorkerClient();
  cleanups.push(() => created.destroy());
  return created;
}

test(`a heap limit on one ${engineKind} worker thread fails only that thread's open`, async () => {
  const bytes = docx();
  const limited = client();
  const unlimited = client();
  const roomy = client();

  const [opened, roomyOpened] = await Promise.all([
    unlimited.openSource({ bytes }),
    roomy.openSource({ bytes, heapLimitBytes: ROOMY_HEAP_LIMIT_BYTES }),
  ]);
  for (const result of [opened, roomyOpened]) {
    expect(result.document.package.document).toBeDefined();
    expect(result.stateVector).toBeInstanceOf(Uint8Array);
    expect(result.stateVector.byteLength).toBeGreaterThan(0);
  }

  const failure = await limited.openSource({ bytes, heapLimitBytes: HEAP_LIMIT_BYTES })
    .then(() => undefined, (error: unknown) => error);
  expect(failure).toBeInstanceOf(Error);
  expect((failure as Error).message).toMatch(OUT_OF_MEMORY);

  const [state, roomyState] = await Promise.all([unlimited.encodeState(), roomy.encodeState()]);
  for (const encoded of [state, roomyState]) expect(encoded.byteLength).toBeGreaterThan(0);

  const exported = await exportDocxMarkdown(bytes, { revisionView: 'accepted' });
  expect(exported.markdown).toContain(LAST_PARAGRAPH);

  await expect(limited.openSource({ bytes })).rejects.toThrow(OUT_OF_MEMORY);

  const session = await createResidentEngineSession();
  cleanups.push(() => session.destroy());
  expect(typeof session.openSourceDocx(bytes, {})).toBe('string');
  expect(session.paragraphs('body').map(paragraph => paragraph.text)).toContain(LAST_PARAGRAPH);
}, 30_000);
