import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import { sessionInternals } from './sessionInternals';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

function docx(): Uint8Array {
  const parts: PartsMap = new Map();
  const set = (name: string, content: string) => parts.set(name, toBytes(content));
  set(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  );
  set(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  );
  set(
    'word/document.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Source bytes</w:t></w:r></w:p></w:body></w:document>'
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function watchSource(session: YrsSession): WeakRef<ArrayBuffer> {
  const source = sessionInternals(session).sourcePackage();
  expect(source).not.toBeNull();
  expect(source!.buffer.byteLength).toBeGreaterThan(0);
  expect(sessionInternals(session).sourcePackage()!.buffer).toBe(source!.buffer);
  return new WeakRef(source!.buffer);
}

async function gcTick(): Promise<void> {
  await new Promise<void>((resolve) =>
    setTimeout(() => {
      Bun.gc(true);
      resolve();
    }, 0)
  );
  await new Promise((resolve) => setTimeout(resolve, 0));
}

async function collect<T extends object>(bytes: WeakRef<T>): Promise<void> {
  for (let i = 0; i < 20; i++) {
    await gcTick();
    if (bytes.deref() === undefined) return;
  }
}

test('a destroyed session releases its source bytes while the facade is retained', async () => {
  const session = await createYrsSession({ clientId: 76201 });
  try {
    session.openDocx(docx(), true);
    const source = watchSource(session);
    await gcTick();
    expect(source.deref() !== undefined).toBe(true);

    session.destroy();
    await collect(source);
    expect(source.deref()).toBeUndefined();
    expect(sessionInternals(session).sourcePackage()).toBeNull();
    expect(() => session.destroy()).not.toThrow();
  } finally {
    session.destroy();
  }
});

test('a destroyed worker mirror releases its registered font bytes while the facade is retained', async () => {
  const session = await createYrsSession({ clientId: 76202 });
  try {
    session.createStory('body', 'Font bytes');
    const bytes = new Uint8Array(readFileSync(FONT));
    expect(session.registerFont(bytes)).toBe(0);
    session.adoptResidentWorkerLayout!(JSON.stringify({
      bodyStory: 'body',
      regions: { sections: [{ sectionId: 'main', properties: {} }] },
      measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
      renderEnv: {},
    }));
    session.mirrorWorkerDocument({
      version: session.version(),
      proposals: { previewVersion: 0, entries: [] },
    });
    const snapshot = session.residentWorkerSnapshot()!;
    expect(snapshot.workerAuthoritative).toBe(true);
    expect(snapshot.fonts).toEqual([bytes]);

    session.destroy();
    expect(session.residentWorkerSnapshot()!.fonts).toEqual([]);
    expect(() => session.destroy()).not.toThrow();
  } finally {
    session.destroy();
  }
});
