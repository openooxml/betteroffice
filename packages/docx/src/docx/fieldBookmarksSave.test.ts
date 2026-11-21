import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '.';
import { preloadEditWasm } from '../wasm/edit';
import { assertBookmarkInvariant } from './__testSupport__/bookmarkInvariant';
import { repackDocx } from './rezip';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { unzipContainer } from './wasm';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';
const separate = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const instruction = (text: string) => `<w:r><w:instrText xml:space="preserve"> ${text} </w:instrText></w:r>`;

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))), 30_000);

function fixture(body: string): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  add('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function repacked(source: Uint8Array): Promise<string> {
  const saved = new Uint8Array(await repackDocx(await parseDocx(source.slice().buffer, { preloadFonts: false }), { updateModifiedDate: false }));
  assertBookmarkInvariant(source, saved);
  return new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
}

function assertOrder(xml: string, fragments: string[]): void {
  let offset = 0;
  for (const fragment of fragments) {
    const at = xml.indexOf(fragment, offset);
    expect(at, `${fragment} in ${xml}`).toBeGreaterThanOrEqual(0);
    offset = at + fragment.length;
  }
}

test('a form field keeps the bookmark Word writes inside its field code', async () => {
  const start = '<w:bookmarkStart w:id="0" w:name="Text1"/>';
  const stop = '<w:bookmarkEnd w:id="0"/>';
  const xml = await repacked(fixture(`<w:p>${begin}${start}${instruction('FORMTEXT')}${separate}<w:r><w:t>Value</w:t></w:r>${end}${stop}</w:p>`));
  assertOrder(xml, ['"begin"', start, 'FORMTEXT', '"separate"', '>Value<', '"end"', stop]);
});

test('a field nested in a field code survives a parse and repack', async () => {
  const nested = `${begin}${instruction('PAGE')}${separate}<w:r><w:t>2</w:t></w:r>${end}`;
  const xml = await repacked(fixture(`<w:p>${begin}${instruction('=')}${nested}${instruction('/2+1')}${separate}<w:r><w:t>2</w:t></w:r>${end}</w:p>`));
  assertOrder(xml, ['"begin"', ' = ', '"begin"', ' PAGE ', '"separate"', '"end"', '/2+1', '"separate"', '"end"']);
});
