import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '.';
import { preloadEditWasm } from '../wasm/edit';
import { repackDocx } from './rezip';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';
import { unzipContainer } from './wasm';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const START = '<w:bookmarkStart w:id="0" w:name="_Toc1"/>';
const END = '<w:bookmarkEnd w:id="0"/>';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))), 30_000);

function fixture(byteOrderMark: boolean): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  add('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/document.xml', `${byteOrderMark ? '﻿' : ''}<w:document xmlns:w="${W}"><w:body>` +
    `${START}<w:p><w:r><w:t>Heading QA</w:t></w:r></w:p>${END}<w:p><w:r><w:t>Body QA</w:t></w:r></w:p>` +
    '</w:body></w:document>');
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

for (const byteOrderMark of [false, true]) {
  test(`body bookmarks survive a parse and repack ${byteOrderMark ? 'with' : 'without'} a byte order mark`, async () => {
    const document = await parseDocx(fixture(byteOrderMark).buffer as ArrayBuffer, { preloadFonts: false });
    const saved = new Uint8Array(await repackDocx(document, { updateModifiedDate: false }));
    const xml = new TextDecoder().decode(unzipContainer(saved)['word/document.xml']);
    const order = [START, 'Heading QA', END, 'Body QA'].map((needle) => xml.indexOf(needle));
    expect(order.every((at) => at >= 0)).toBe(true);
    expect(order).toEqual([...order].sort((left, right) => left - right));
    await parseDocx(saved.buffer as ArrayBuffer, { preloadFonts: false });
  });
}
