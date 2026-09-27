import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef, useState, type RefObject } from 'react';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import {
  createDisplayListQueries,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import type { Document } from '@betteroffice/docx/types/document';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { preloadLayoutWasm } from '@betteroffice/docx/wasm/layout';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { PagedEditor, type PagedEditorRef } from './PagedEditor';
import { useYrsCoreSession, type YrsCoreSession } from './hooks/useYrsCoreSession';
import { markPresented, stampSourceVersion } from './internals/layoutProvenance';
import { createYrsPositionProjection } from './internals/yrsPositionProjection';

const { act, cleanup, fireEvent, render } = await import('@testing-library/react');

const GENERATED = resolve(import.meta.dir, '../../../../docx/src/wasm/generated');
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WORD = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const TEXT = 'Drop the value here';
const disposables: DisplayListQueries[] = [];

let fontBytes: ArrayBuffer;

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {} },
      configurable: true,
    });
  }
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(GENERATED, 'edit/docx_edit_bg.wasm'))));
  await preloadLayoutWasm(
    new Uint8Array(readFileSync(resolve(GENERATED, 'layout/docx_layout_bg.wasm')))
  );
  fontBytes = readFileSync(FONT).buffer as ArrayBuffer;
});
afterEach(() => {
  cleanup();
  for (const queries of disposables.splice(0)) queries.dispose();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function fixture(): Uint8Array {
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
    toBytes(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>${TEXT}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
    )
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

interface Painted {
  queries: DisplayListQueries;
  hostRef: RefObject<HTMLDivElement | null>;
}

function Harness({
  bytes,
  editorRef,
  coreRef,
  painted,
}: {
  bytes: Uint8Array;
  editorRef: RefObject<PagedEditorRef | null>;
  coreRef: { current: YrsCoreSession | null };
  painted: Painted | null;
}) {
  const [host, setHost] = useState<Document | null>(null);
  const core = useYrsCoreSession(true, host, null, bytes, 0, undefined, {
    onHostDocument: (next) => setHost(next.document),
  });
  coreRef.current = host ? core : null;
  return (
    <PagedEditor
      ref={editorRef}
      document={host}
      yrsCore={core}
      measurementFontProvider={{ resolve: () => () => Promise.resolve(fontBytes) }}
      displayListQueries={painted?.queries ?? null}
      canvasHostRef={painted?.hostRef}
    />
  );
}

async function until(done: () => boolean): Promise<void> {
  const deadline = Date.now() + 5000;
  while (!done() && Date.now() < deadline) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
  }
  expect(done()).toBe(true);
}

/** Canvas pages that finished painting the session's current version. */
async function paint(session: YrsSession): Promise<Painted> {
  const start = createYrsPositionProjection(session, 'body')!.positionForLoc({
    story: 'body',
    paraId: '00000001',
    offset: 0,
  })!;
  const queries = createDisplayListQueries({
    pages: [
      {
        pageIndex: 0,
        width: 800,
        height: 1000,
        contentBounds: { x: 80, y: 80, width: 640, height: 760 },
        primitives: [
          {
            kind: 'text',
            text: TEXT,
            x: 100,
            baselineY: 150,
            width: 200,
            font: '400 16px Calibri',
            color: '#000000',
            docStart: start,
            docEnd: start + TEXT.length,
            blockId: start,
            lineIndex: 0,
          },
        ],
      },
    ],
  });
  await queries.whenReady();
  disposables.push(queries);
  stampSourceVersion(queries, session.version());
  const host = document.createElement('div');
  host.className = 'canvas-pages';
  const page = document.createElement('div');
  page.className = 'canvas-page';
  page.dataset.pageIndex = '0';
  const canvas = document.createElement('canvas');
  canvas.dataset.pageIndex = '0';
  canvas.getBoundingClientRect = () => new DOMRect(0, 0, 800, 1000);
  page.getBoundingClientRect = canvas.getBoundingClientRect;
  page.append(canvas);
  host.append(page);
  markPresented(host, queries.displayList);
  return { queries, hostRef: { current: host } };
}

test('a point query waits for composed input to reach the document', async () => {
  const editorRef = createRef<PagedEditorRef>();
  const coreRef: { current: YrsCoreSession | null } = { current: null };
  const bytes = fixture();
  const view = render(
    <Harness bytes={bytes} editorRef={editorRef} coreRef={coreRef} painted={null} />
  );
  await until(() => !!coreRef.current?.session && !!editorRef.current?.getLayout());
  const session = editorRef.current!.getYrsSession()!;
  const painted = await paint(session);
  view.rerender(
    <Harness bytes={bytes} editorRef={editorRef} coreRef={coreRef} painted={painted} />
  );
  const at = () => editorRef.current!.getPositionAtPoint(210, 145);
  expect(at()).toMatchObject({ version: session.version(), target: { story: 'body' } });

  const textarea = view.getByTestId('yrs-input') as HTMLTextAreaElement;
  fireEvent.compositionStart(textarea);
  expect(editorRef.current!.hasPendingInput()).toBe(true);
  expect(at()).toBeNull();

  textarea.value = '日';
  fireEvent.compositionEnd(textarea, { data: '日' });
  expect(at()).toBeNull();
  await act(() => editorRef.current!.flushPendingInput());
  expect(editorRef.current!.hasPendingInput()).toBe(false);
  expect(session.readParagraphs({ view: 'accepted' })).toMatchObject({
    paragraphs: [{ text: expect.stringContaining('日') }],
  });
  expect(at()).toBeNull();
  const repainted = await paint(session);
  view.rerender(
    <Harness bytes={bytes} editorRef={editorRef} coreRef={coreRef} painted={repainted} />
  );
  expect(at()).toMatchObject({ version: session.version() });
});
