import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer } from '../docx/rezip/parts';
import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, docxPackageDigest } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
let DOCX: Uint8Array;

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  const parts = unzipContainer(
    new Uint8Array(
      readFileSync(
        resolve(
          import.meta.dir,
          '../../../../crates/docx-edit/tests/fixtures/suppressed-list-markers.docx'
        )
      )
    )
  );
  // An invalid text ID, which parsing replaces with one the package digest seeds.
  const document = new TextDecoder()
    .decode(parts['word/document.xml'])
    .replace('<w:p>', '<w:p w14:textId="invalid">')
    .replace(
      '<w:document ',
      '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    );
  expect(document).toContain('w14:textId="invalid"');
  parts['word/document.xml'] = new TextEncoder().encode(document);
  DOCX = new Uint8Array(rezipPartsToArrayBuffer(new Map(Object.entries(parts))));
});

async function opened(digest?: string): Promise<{ state: Uint8Array; package: string }> {
  const session = await createYrsSession({ clientId: 91 });
  try {
    const host = session.openDocx(DOCX, true, {
      generation: 'digest',
      ...(digest ? { digest } : {}),
    });
    return {
      state: session.encodeState(),
      package: JSON.stringify([host.document, session.materializeDocx()]),
    };
  } finally {
    session.destroy();
  }
}

test('an open given the Web Crypto digest matches an open that hashes the package', async () => {
  const digest = await docxPackageDigest(DOCX);
  expect(digest).toMatch(/^[0-9a-f]{64}$/);
  const hashed = await opened();
  const given = await opened(digest);
  expect(Buffer.from(given.state).equals(Buffer.from(hashed.state))).toBe(true);
  expect(given.package).toBe(hashed.package);
  expect((await opened('0'.repeat(64))).package).not.toBe(hashed.package);
});

test('an open refuses a digest that is not one', async () => {
  const session = await createYrsSession({ clientId: 92 });
  try {
    expect(() => session.openDocx(DOCX, true, { digest: 'ABC' })).toThrow();
  } finally {
    session.destroy();
  }
});
