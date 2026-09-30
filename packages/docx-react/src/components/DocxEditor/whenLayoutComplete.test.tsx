import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { parseDocx } from '@betteroffice/docx/docx';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';

const { act, cleanup, render } = await import('@testing-library/react');
const { DocxEditor } = await import('../../index');
type DocxEditorRef = import('../../index').DocxEditorRef;

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const quiet = { error: console.error, warn: console.warn };

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: {
        addEventListener: () => {},
        removeEventListener: () => {},
        ready: Promise.resolve(),
      },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  console.error = () => {};
  console.warn = () => {};
});
afterEach(cleanup);
afterAll(async () => {
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function pagedDocx(pages: number): Promise<ArrayBuffer> {
  const body = Array.from(
    { length: pages },
    (_, index) =>
      `<w:p>${index ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : ''}<w:r><w:t>Page ${index + 1}</w:t></w:r></w:p>`
  ).join('');
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
  );
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function tick(ms = 10) {
  await act(async () => {
    await new Promise((done) => setTimeout(done, ms));
  });
}

/** The page count `whenLayoutComplete()` resolves with, or its failure. */
async function layoutComplete(ref: React.RefObject<DocxEditorRef | null>) {
  let outcome = null as number | Error | null;
  ref.current!.whenLayoutComplete().then(
    (pages) => (outcome = pages),
    (error: Error) => (outcome = error)
  );
  for (let attempt = 0; attempt < 300 && outcome === null; attempt += 1) await tick();
  return outcome;
}

async function mountTwoPages() {
  const ref = createRef<DocxEditorRef>();
  render(<DocxEditor ref={ref} documentBuffer={await pagedDocx(2)} />);
  for (let attempt = 0; attempt < 300 && !ref.current; attempt += 1) await tick();
  expect(await layoutComplete(ref)).toBe(2);
  return ref;
}

test('a parsed reload settles with the new document', async () => {
  const ref = await mountTwoPages();
  const document = await parseDocx(await pagedDocx(3));
  act(() => ref.current!.loadDocument(document));
  expect(await layoutComplete(ref)).toBe(3);
  expect(ref.current!.getTotalPages()).toBe(3);
}, 30_000);

test('a parsed reload with the same page count reports it again', async () => {
  const ref = await mountTwoPages();
  const document = ref.current!.getDocument()!;
  act(() => ref.current!.loadDocument(document));
  expect(await layoutComplete(ref)).toBe(2);
  expect(ref.current!.getTotalPages()).toBe(2);
}, 30_000);

test('a reload waits for the new document and reports no pages while it loads', async () => {
  const ref = await mountTwoPages();

  let release = () => {};
  const gate = new Promise<void>((done) => (release = done));
  const bytes = await pagedDocx(3);
  const blob = new Blob([bytes]);
  blob.arrayBuffer = async () => {
    await gate;
    return bytes;
  };
  let loading: Promise<void> = Promise.resolve();
  await act(async () => {
    loading = ref.current!.loadDocumentBuffer(blob);
  });
  let next = null as number | null;
  void ref.current!.whenLayoutComplete().then((pages) => (next = pages));
  await tick(200);
  expect(next).toBeNull();
  expect(ref.current!.getTotalPages()).toBe(0);

  release();
  for (let attempt = 0; attempt < 300 && next === null; attempt += 1) await tick();
  await loading;
  expect(next).toBe(3);
  expect(ref.current!.getTotalPages()).toBe(3);
}, 30_000);

test('each failed load rejects the wait', async () => {
  const ref = await mountTwoPages();
  const detached = await pagedDocx(1);
  structuredClone(detached, { transfer: [detached] });
  for (let load = 0; load < 2; load += 1) {
    await act(async () => {
      await ref.current!.loadDocumentBuffer(detached);
    });
    expect(await layoutComplete(ref)).toBeInstanceOf(Error);
  }
}, 30_000);

test('a failed load before any layout rejects the wait', async () => {
  const ref = createRef<DocxEditorRef>();
  render(<DocxEditor ref={ref} />);
  for (let attempt = 0; attempt < 300 && !ref.current; attempt += 1) await tick();
  const detached = await pagedDocx(1);
  structuredClone(detached, { transfer: [detached] });
  await act(async () => {
    await ref.current!.loadDocumentBuffer(detached);
  });
  await tick(200);
  expect(await layoutComplete(ref)).toBeInstanceOf(Error);
}, 30_000);

test('a failure without a message rejects waits during and after the load', async () => {
  const ref = await mountTwoPages();
  let fail = () => {};
  const blob = new Blob([]);
  blob.arrayBuffer = () => new Promise((_, reject) => (fail = () => reject(new Error())));
  await act(async () => {
    void ref.current!.loadDocumentBuffer(blob);
  });
  const during = layoutComplete(ref);
  fail();
  expect(await during).toBeInstanceOf(Error);
  expect(await layoutComplete(ref)).toBeInstanceOf(Error);
}, 30_000);
