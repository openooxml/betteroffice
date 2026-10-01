import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { createRef, useState, type RefObject } from 'react';
import { repackDocx } from '@betteroffice/docx/docx';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { unzipContainer } from '@betteroffice/docx/docx/wasm';
import type { Document } from '@betteroffice/docx/types/document';
import { sessionInternals } from '@betteroffice/docx/yrs/sessionInternals';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { partEditStory } from './partEdit';
import { PagedEditor, type PagedEditorRef } from './PagedEditor';
import type { PagedEditorCommandBridge } from './hooks/usePagedEditorRefApi';
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

function fixture(sharedHeader = false): Uint8Array {
  const parts = new Map<string, Uint8Array>();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${WORD}.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="${WORD}.header+xml"/></Types>`
    )
  );
  parts.set(
    '_rels/.rels',
    toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    )
  );
  parts.set(
    'word/_rels/document.xml.rels',
    toBytes(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/>${sharedHeader ? `<Relationship Id="rIdHeaderAlias" Type="${R}/header" Target="./header1.xml"/>` : ''}</Relationships>`
    )
  );
  parts.set(
    'word/document.xml',
    toBytes(
      `<w:document ${NS}><w:body><w:p w14:paraId="00000001">${sharedHeader ? '<w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr></w:pPr>' : ''}<w:r><w:t>Body text</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="${sharedHeader ? 'rIdHeaderAlias' : 'rIdHeader'}"/></w:sectPr></w:body></w:document>`
    )
  );
  parts.set(
    'word/header1.xml',
    toBytes(`<w:hdr ${NS}><w:p w14:paraId="0000E001"><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`)
  );
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function Harness({
  bytes,
  editorRef,
  bridgeRef,
  coreRef,
}: {
  bytes: Uint8Array;
  editorRef: RefObject<PagedEditorRef | null>;
  bridgeRef: { current: PagedEditorCommandBridge | null };
  coreRef: { current: YrsCoreSession | null };
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

/** The body and header texts a save of the editor's current document writes. */
async function saved(editorRef: RefObject<PagedEditorRef | null>): Promise<string[]> {
  const document = editorRef.current?.getDocument();
  if (!document) throw new Error('the editor has no document to save');
  const parts = unzipContainer(new Uint8Array(await act(() => repackDocx(document))));
  return ['word/document.xml', 'word/header1.xml'].map((name) =>
    [...new TextDecoder().decode(parts[name] as Uint8Array).matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)]
      .map((match) => match[1])
      .join('')
  );
}

test('undo and redo refresh every story a multi-story batch changed before the next save', async () => {
  const editorRef = createRef<PagedEditorRef>();
  const bridgeRef: { current: PagedEditorCommandBridge | null } = { current: null };
  const coreRef: { current: YrsCoreSession | null } = { current: null };
  render(<Harness bytes={fixture()} editorRef={editorRef} bridgeRef={bridgeRef} coreRef={coreRef} />);
  await until(
    () => !!coreRef.current?.session && !!editorRef.current?.getYrsSession() && !!bridgeRef.current
  );
  const session = coreRef.current!.session!;
  const result = session.applyEdits({
    expectVersion: session.version(),
    steps: [
      { op: 'replaceText', target: { kind: 'paragraph', story: 'body', paraId: '00000001' }, text: 'Body edited' },
      {
        op: 'replaceText',
        target: { kind: 'paragraph', story: 'hf:rIdHeader', paraId: '0000E001' },
        text: 'Header edited',
      },
    ],
  });
  if (!result.ok) throw new Error(result.failure.message);
  act(() => {
    editorRef.current!.syncYrsInputState(true, result.changedStories);
  });
  expect(await saved(editorRef)).toEqual(['Body edited', 'Header edited']);

  act(() => {
    expect(editorRef.current!.undo()).toBe(true);
  });
  expect(await saved(editorRef)).toEqual(['Body text', 'Header text']);

  const bridge = bridgeRef.current!;
  expect(await act(() => bridge.runAfterPendingInput(() => bridge.history(true)))).toBe(true);
  expect(await saved(editorRef)).toEqual(['Body edited', 'Header edited']);
});

test('default aliased package keeps the second section header story', async () => {
  const editorRef = createRef<PagedEditorRef>();
  const bridgeRef: { current: PagedEditorCommandBridge | null } = { current: null };
  const coreRef: { current: YrsCoreSession | null } = { current: null };
  render(<Harness bytes={fixture(true)} editorRef={editorRef} bridgeRef={bridgeRef} coreRef={coreRef} />);
  await until(() => !!coreRef.current?.session && !!editorRef.current?.getYrsSession());
  const session = coreRef.current!.session!;
  expect(session.hasStory('hf:rIdHeader')).toBe(true);
  expect(session.hasStory('hf:rIdHeaderAlias')).toBe(true);
  expect(partEditStory({ kind: 'header', rId: 'rIdHeaderAlias' }, session))
    .toBe('hf:rIdHeaderAlias');
  expect([...sessionInternals(session).headerFooterAliases()]).toEqual([]);
});

test('editing the second section header projects and saves canonical content for both rIds', async () => {
  const editorRef = createRef<PagedEditorRef>();
  const bridgeRef: { current: PagedEditorCommandBridge | null } = { current: null };
  const coreRef: { current: YrsCoreSession | null } = { current: null };
  render(<Harness bytes={fixture(true)} editorRef={editorRef} bridgeRef={bridgeRef} coreRef={coreRef} />);
  await until(() => !!coreRef.current?.session && !!editorRef.current?.getYrsSession());
  const session = coreRef.current!.session!;
  act(() => {
    for (const story of session.storyIds()) {
      if (story === 'hf:rIdHeaderAlias' || story.startsWith('hf:rIdHeaderAlias:')) {
        session.deleteStory(story);
      }
    }
    sessionInternals(session).setHeaderFooterAliases(JSON.stringify([{
      isHeader: true, partPath: 'word/header1.xml',
      relationshipIds: ['rIdHeader', 'rIdHeaderAlias'],
    }]));
    editorRef.current!.syncYrsInputState(true, ['hf:rIdHeader', 'hf:rIdHeaderAlias']);
  });
  const part = { kind: 'header', rId: 'rIdHeaderAlias' } as const;
  const story = partEditStory(part, session);
  expect(part.rId).toBe('rIdHeaderAlias');
  expect(story).toBe('hf:rIdHeader');
  expect(session.hasStory('hf:rIdHeaderAlias')).toBe(false);
  const paragraph = session.paragraphs(story)[0]!;
  act(() => {
    session.insertText({ story, paraId: paragraph.paraId, offset: 11 }, ' edited');
    editorRef.current!.syncYrsInputState(true, [story]);
  });
  const document = editorRef.current!.getDocument()!;
  expect(document.package.headers!.get('rIdHeader')!.content)
    .toBe(document.package.headers!.get('rIdHeaderAlias')!.content);
  expect(await saved(editorRef)).toEqual(['Body text', 'Header text edited']);
});
