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
import type { YrsSession, YrsStoryRange } from '@betteroffice/docx/yrs';
import { PagedEditor, type PagedEditorRef } from './PagedEditor';
import type { PagedEditorCommandBridge } from './hooks/usePagedEditorRefApi';
import { useYrsCoreSession, type YrsCoreSession } from './hooks/useYrsCoreSession';
import { yrsCellStory } from './yrsCommands';

const { act, cleanup, render } = await import('@testing-library/react');

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(import.meta.dir, '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

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
}, 30_000);
afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const run = (text: string, size?: number) =>
  `<w:r>${size ? `<w:rPr><w:sz w:val="${size}"/></w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const cell = (runs: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p>${runs}</w:p></w:tc>`;

/** Body text above a 2x2 table: Alpha at 10pt, Beta at 14pt, Gamma unsized, Delta half at 14pt. */
function fixture(): Uint8Array {
  const table =
    '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell(run('Alpha', 20))}${cell(run('Beta', 28))}</w:tr>` +
    `<w:tr>${cell(run('Gamma'))}${cell(run('Del', 28) + run('ta'))}</w:tr></w:tbl>`;
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
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
    toBytes(`<w:document xmlns:w="${W}"><w:body><w:p>${run('Body text')}</w:p>${table}<w:p>${run('After')}</w:p><w:sectPr/></w:body></w:document>`)
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function Harness({
  editorRef,
  bridgeRef,
  coreRef,
}: {
  editorRef: RefObject<PagedEditorRef | null>;
  bridgeRef: { current: PagedEditorCommandBridge | null };
  coreRef: { current: YrsCoreSession | null };
}) {
  const [host, setHost] = useState<Document | null>(null);
  const [bytes] = useState(fixture);
  const core = useYrsCoreSession(true, host, null, bytes, 0, undefined, {
    onHostDocument: (next) => setHost(next.document),
  });
  coreRef.current = host ? core : null;
  return (
    <PagedEditor
      ref={editorRef}
      document={host}
      yrsCore={core}
      commandBridgeRef={bridgeRef}
      measurementFontProvider={{ resolve: () => () => Promise.resolve(fontBytes) }}
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

function whole(session: YrsSession, story: string): YrsStoryRange {
  const spans = session.paragraphSpans(story);
  const last = spans[spans.length - 1]!;
  return { story, start: { paraId: spans[0]!.paraId, offset: 0 }, end: { paraId: last.paraId, offset: last.length } };
}

async function opened() {
  const editorRef = createRef<PagedEditorRef>();
  const bridgeRef: { current: PagedEditorCommandBridge | null } = { current: null };
  const coreRef: { current: YrsCoreSession | null } = { current: null };
  render(<Harness editorRef={editorRef} bridgeRef={bridgeRef} coreRef={coreRef} />);
  await until(() => !!coreRef.current?.session && !!editorRef.current?.getYrsSession() && !!bridgeRef.current);
  const session = coreRef.current!.session!;
  const table = { story: 'body', tableIndex: 0 };
  const cells = [[0, 0], [0, 1], [1, 0], [1, 1]].map(([row, column]) => yrsCellStory(session, { ...table, row: row!, column: column! })!);
  const body = session.paragraphs('body').find((paragraph) => paragraph.text === 'Body text')!.paraId;
  const bodyText: YrsStoryRange = { story: 'body', start: { paraId: body, offset: 0 }, end: { paraId: body, offset: 9 } };
  const caretIn = (story: string) => {
    act(() => {
      session.setSelection({ story, paraId: session.paragraphs(story)[0]!.paraId, offset: 0 });
      editorRef.current!.syncYrsInputState(false);
    });
  };
  return {
    session,
    bridge: bridgeRef.current!,
    editor: () => editorRef.current!,
    table,
    cells,
    bodyText,
    caretIn,
    sizes: () => cells.map((story) => session.selectionContext(whole(session, story)).fontSize),
    marks: (mark: 'bold' | 'italic') => cells.map((story) => session.selectionContext(whole(session, story))[mark]),
  };
}

test('select table, then a toolbar size and bold, format every cell and undo one action at a time', async () => {
  const { session, bridge, editor, cells, bodyText, caretIn, sizes, marks } = await opened();
  expect(sizes()).toEqual([20, 28, null, null]);
  caretIn(cells[0]!);
  act(() => {
    expect(bridge.command({ type: 'tableSelect', target: 'table' })).toBe(true);
  });
  expect(session.cellSelection()).toMatchObject({ anchor: { row: 0, column: 0 }, head: { row: 1, column: 1 } });

  act(() => {
    expect(bridge.format({ type: 'fontSize', value: 12 })).toBe(true);
  });
  expect(sizes()).toEqual([24, 24, 24, 24]);
  act(() => {
    expect(bridge.format('bold')).toBe(true);
  });
  expect(marks('bold')).toEqual([true, true, true, true]);
  expect(session.selectionContext(bodyText)).toMatchObject({ bold: false, fontSize: null });

  act(() => {
    expect(editor().undo()).toBe(true);
  });
  expect(marks('bold')).toEqual([false, false, false, false]);
  expect(sizes()).toEqual([24, 24, 24, 24]);
  act(() => {
    expect(editor().undo()).toBe(true);
  });
  expect(sizes()).toEqual([20, 28, null, null]);
});

test('a dragged range of cells takes the toolbar formatting on each of its cells only', async () => {
  const { session, bridge, table, cells, caretIn, sizes, marks } = await opened();
  caretIn(cells[1]!);
  act(() => {
    session.setCellSelection({ anchor: { ...table, row: 0, column: 1 }, head: { ...table, row: 1, column: 1 } });
    expect(bridge.format('italic')).toBe(true);
    expect(bridge.format({ type: 'fontSize', value: 9 })).toBe(true);
  });
  expect(marks('italic')).toEqual([false, true, false, true]);
  expect(sizes()).toEqual([20, 18, null, 18]);
});

test('a caret without a cell range still stores the formatting for the next typed text', async () => {
  const { session, bridge, editor, cells, bodyText, caretIn, marks } = await opened();
  act(() => {
    session.setSelection({ story: 'body', ...bodyText.start, offset: 4 });
    editor().syncYrsInputState(false);
    expect(bridge.format('bold')).toBe(true);
  });
  expect(session.selectionContext(bodyText).bold).toBe(false);
  expect(bridge.toolbarSelection(false)?.context.bold).toBe(true);

  caretIn(cells[0]!);
  act(() => {
    expect(bridge.format('italic')).toBe(true);
  });
  expect(marks('italic')).toEqual([false, false, false, false]);
  expect(bridge.toolbarSelection(false)?.context.italic).toBe(true);
});
