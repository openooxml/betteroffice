import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '.';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';

let wasm = true;
try {
  await import('./rustParseFacade');
} catch {
  wasm = false;
}
const describeIfWasm = wasm ? describe : describe.skip;

const README = readFileSync(resolve(import.meta.dir, '../../README.md'), 'utf8');

function minimalDocx(): ArrayBuffer {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>'
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rIdPkg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>'
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>'
    )
  );
  return rezipPartsToArrayBuffer(parts);
}

describeIfWasm('README', () => {
  test('every model path the parse example names exists on a parsed document', async () => {
    const paths = [...new Set(README.match(/\bdocument\.package(?:\.[A-Za-z]+)*/g) ?? [])];
    expect(paths).toContain('document.package.document.content');
    const document = await parseDocx(minimalDocx());
    for (const path of paths) {
      const value = path
        .split('.')
        .slice(1)
        .reduce<unknown>(
          (node, key) => (node as Record<string, unknown> | undefined)?.[key],
          document
        );
      expect({ path, defined: value !== undefined }).toEqual({ path, defined: true });
    }
  });
});
