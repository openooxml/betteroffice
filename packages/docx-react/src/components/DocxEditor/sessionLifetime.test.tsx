import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, expect, mock, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';

const { act, cleanup, render } = await import('@testing-library/react');

// Records every call a session gets after it was destroyed.
const real = await import('@betteroffice/docx/yrs');
const { createYrsSession } = real;
const afterDestroy: string[] = [];
const live = new Set<unknown>();
mock.module('@betteroffice/docx/yrs', () => ({
  ...real,
  createYrsSession: async (options: Parameters<typeof createYrsSession>[0]) => {
    const session = await createYrsSession(options);
    let destroyed = false;
    const proxy: typeof session = new Proxy(session, {
      get(target, key, receiver) {
        const value = Reflect.get(target, key, receiver);
        if (typeof value !== 'function') return value;
        return (...args: unknown[]) => {
          if (key === 'destroy') {
            destroyed = true;
            live.delete(proxy);
          } else if (destroyed) {
            afterDestroy.push(String(key));
          }
          return value.apply(target, args);
        };
      },
    });
    live.add(proxy);
    return proxy;
  },
}));
const { DocxEditor } = await import('../../index');

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const quiet = { error: console.error, warn: console.warn };

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {}, ready: Promise.resolve() },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  console.error = () => {};
  console.warn = () => {};
});
afterAll(async () => {
  cleanup();
  await new Promise((done) => setTimeout(done, 100));
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

/** A document of `pages` pages, each with a commented paragraph. */
async function commentedDocx(pages: number): Promise<ArrayBuffer> {
  const body = Array.from(
    { length: pages },
    (_, i) =>
      `<w:p>${i ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : ''}<w:commentRangeStart w:id="${i}"/>` +
      `<w:r><w:t>Page ${i + 1}</w:t></w:r><w:commentRangeEnd w:id="${i}"/>` +
      `<w:r><w:commentReference w:id="${i}"/></w:r></w:p>`
  ).join('');
  const comments = Array.from(
    { length: pages },
    (_, i) =>
      `<w:comment w:id="${i}" w:author="A" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Note ${i}</w:t></w:r></w:p></w:comment>`
  ).join('');
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
  );
  zip.file(
    'word/_rels/document.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
      '</Relationships>'
  );
  zip.file('word/document.xml', `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`);
  zip.file('word/comments.xml', `<w:comments xmlns:w="${W}">${comments}</w:comments>`);
  return zip.generateAsync({ type: 'arraybuffer' });
}

async function settle(ms: number) {
  for (let waited = 0; waited < ms; waited += 50) {
    await act(async () => {
      await new Promise((done) => setTimeout(done, 50));
    });
  }
}

test('loading another document never calls into the session it replaced', async () => {
  const view = render(<DocxEditor documentBuffer={await commentedDocx(3)} />);
  await settle(2500);
  view.rerender(<DocxEditor documentBuffer={await commentedDocx(4)} />);
  await settle(3000);
  expect(afterDestroy).toEqual([]);
  view.unmount();
}, 30_000);

test('a load that fails frees the document it replaced', async () => {
  afterDestroy.length = 0;
  const view = render(<DocxEditor documentBuffer={await commentedDocx(2)} />);
  await settle(2500);
  const detached = await commentedDocx(1);
  structuredClone(detached, { transfer: [detached] });
  view.rerender(<DocxEditor documentBuffer={detached} />);
  await settle(1500);
  expect(live.size).toBe(0);
  expect(afterDestroy).toEqual([]);
  view.unmount();
}, 30_000);
