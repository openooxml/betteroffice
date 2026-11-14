import { beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import { preloadEditWasm } from '../wasm/edit';
import { residentWorkerFactory } from './__fixtures__/residentWorker';
import { ResidentEngineWorkerClient } from './residentEngineWorkerClient';

const RUN = '<w:r><w:rPr><w:w w:val="0"/></w:rPr><w:t>Scaled text</w:t></w:r>';

function fixture(): Uint8Array<ArrayBuffer> {
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([
    ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>${RUN}</w:p><w:p><w:r><w:t>Plain text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`)],
  ])));
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))), 30_000);

it('a save keeps a run scale of 0 as the file wrote it', async () => {
  const startWorker = await residentWorkerFactory();
  const client = new ResidentEngineWorkerClient(startWorker(74023));
  try {
    await client.open(fixture());
    const saved = readDocxContainer((await client.save({ comments: [] })).bytes).text('word/document.xml') ?? '';
    expect(saved).toContain(RUN);
    expect(saved).not.toContain('w:val="100"');
  } finally {
    client.destroy();
  }
});
