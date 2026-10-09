import { afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { rezipPartsToArrayBuffer } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { unzipContainer } from '../wasm/opc';
import { createYrsSession, prepareDocxBytes, saveYrsDocx } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const crypto = globalThis.crypto;
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

afterEach(() => {
  Object.defineProperty(globalThis, 'crypto', { value: crypto, configurable: true });
});

/** What opening `bytes` gives: its state, package and the save of one edit. */
async function opened(bytes: Uint8Array) {
  const session = await createYrsSession({ clientId: 91 });
  try {
    const host = session.openDocx(bytes, true, { generation: 'digest' });
    const before = { state: session.encodeState(), package: session.materializeDocx() };
    const paragraph = session.paragraphs('body')[0]!;
    session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Edited ');
    const saved = await saveYrsDocx(session);
    return {
      state: Buffer.from(before.state).toString('base64'),
      package: JSON.stringify([host.document, before.package]),
      saved: Buffer.from(saved.bytes).toString('base64'),
    };
  } finally {
    session.destroy();
  }
}

test('a prepared copy opens, edits and saves as the bytes it copies', async () => {
  const prepared = await prepareDocxBytes(DOCX);
  expect(prepared).not.toBe(DOCX);
  expect(await opened(prepared)).toEqual(await opened(DOCX));
});

test('an open uses the digest prepared with the copy', async () => {
  const prepared = await prepareDocxBytes(DOCX);
  // A zip entry's local modification time: the package parses alike, but hashes differently.
  const changed = prepared.slice();
  changed[10] ^= 1;
  // The copy is the caller's: an open trusts the digest taken of it.
  prepared.set(changed);
  expect((await opened(prepared)).package).not.toBe((await opened(changed)).package);
  expect((await opened(prepared)).package).toBe((await opened(DOCX)).package);
});

test('the bytes prepared are copied before hashing', async () => {
  const source = DOCX.slice();
  const pending = prepareDocxBytes(source);
  source[10] ^= 1;
  expect(await opened(await pending)).toEqual(await opened(DOCX));
});

test('without Web Crypto a prepared copy opens by hashing it', async () => {
  Object.defineProperty(globalThis, 'crypto', { value: undefined, configurable: true });
  const prepared = await prepareDocxBytes(DOCX);
  Object.defineProperty(globalThis, 'crypto', { value: crypto, configurable: true });
  prepared[10] ^= 1;
  expect((await opened(prepared)).package).toBe((await opened(prepared.slice())).package);
});
