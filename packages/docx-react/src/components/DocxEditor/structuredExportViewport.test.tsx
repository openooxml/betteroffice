import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef, useState, type RefObject } from 'react';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import type { Document } from '@betteroffice/docx/types/document';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { PagedEditor, type PagedEditorRef } from './PagedEditor';
import { useYrsCoreSession, type YrsCoreSession } from './hooks/useYrsCoreSession';

const { act, cleanup, render } = await import('@testing-library/react');

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WORD = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"`;

let fontBytes: ArrayBuffer;

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {} },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(WASM)));
  fontBytes = readFileSync(FONT).buffer as ArrayBuffer;
});
afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function fixture(): Uint8Array {
  const paragraphs = Array.from(
    { length: 40 },
    (_, index) =>
      `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0').toUpperCase()}"><w:r><w:t>Paragraph ${index + 1}</w:t></w:r></w:p>`
  ).join('');
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${WORD}.document.main+xml"/></Types>`
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(`<w:document ${NS}><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`)
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function Harness({
  bytes,
  editorRef,
  scrollRef,
  coreRef,
}: {
  bytes: Uint8Array;
  editorRef: RefObject<PagedEditorRef | null>;
  scrollRef: RefObject<HTMLDivElement | null>;
  coreRef: { current: YrsCoreSession | null };
}) {
  const [host, setHost] = useState<Document | null>(null);
  const core = useYrsCoreSession(true, host, null, bytes, 0, undefined, {
    onHostDocument: (next) => setHost(next.document),
  });
  coreRef.current = host ? core : null;
  return (
    <div ref={scrollRef} style={{ height: 400, overflow: 'auto' }}>
      <PagedEditor
        ref={editorRef}
        document={host}
        yrsCore={core}
        scrollContainerRef={scrollRef}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(fontBytes) }}
      />
    </div>
  );
}

async function wait(ms: number): Promise<void> {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, ms));
  });
}

async function until(done: () => boolean): Promise<void> {
  const deadline = Date.now() + 5000;
  while (!done() && Date.now() < deadline) await wait(10);
  expect(done()).toBe(true);
}

/** Waits until the editor's DOM has gone a quiet period without a mutation. */
async function settled(root: Node): Promise<void> {
  let last = Date.now();
  const observer = new MutationObserver(() => {
    last = Date.now();
  });
  observer.observe(root, { subtree: true, childList: true, attributes: true, characterData: true });
  const deadline = Date.now() + 5000;
  while (Date.now() - last < 200 && Date.now() < deadline) await wait(20);
  observer.disconnect();
  expect(Date.now() - last).toBeGreaterThanOrEqual(200);
}

test('exports leave the mounted editor viewport, layout, DOM and selection untouched', async () => {
  const editorRef = createRef<PagedEditorRef>();
  const scrollRef = createRef<HTMLDivElement>();
  const coreRef: { current: YrsCoreSession | null } = { current: null };
  const { container } = render(
    <Harness bytes={fixture()} editorRef={editorRef} scrollRef={scrollRef} coreRef={coreRef} />
  );
  await until(() => !!coreRef.current?.session && !!editorRef.current?.getYrsSession());
  await settled(container);
  const scroller = scrollRef.current!;
  scroller.scrollTop = 240;
  scroller.scrollLeft = 16;
  await settled(container);
  const viewport = [scroller.scrollTop, scroller.scrollLeft];
  expect(viewport).toEqual([240, 16]);
  const editor = editorRef.current!;
  const session = editor.getYrsSession()!;
  const selection = [session.selection(), editor.getSelectionRange()];
  const layout = editor.getLayout();
  expect(layout).not.toBeNull();
  const mutations: MutationRecord[] = [];
  const observer = new MutationObserver((records) => mutations.push(...records));
  observer.observe(container, { subtree: true, childList: true, attributes: true, characterData: true });
  await act(async () => {
    for (const revisionView of ['accepted', 'original', 'markup'] as const) {
      expect(session.exportStructured({ revisionView }).ok).toBe(true);
      expect(session.exportMarkdown({ revisionView, maxBytes: 1_024 }).ok).toBe(true);
    }
  });
  await wait(100);
  observer.disconnect();
  expect(mutations).toHaveLength(0);
  expect(editor.getLayout()).toBe(layout);
  expect([scroller.scrollTop, scroller.scrollLeft]).toEqual(viewport);
  expect([session.selection(), editor.getSelectionRange()]).toEqual(selection);
});
