import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Document } from '@betteroffice/docx/types/document';
import { createFontLoadScope } from '@betteroffice/docx/utils';
import type { FontOption } from '@betteroffice/docx/utils/fontOptions';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsDocxHost } from '@betteroffice/docx/yrs';
import { useHistory } from '../../../hooks/useHistory';
import { createCommentIdAllocator } from '../commentFactories';
import { useDocumentLoader } from './useDocumentLoader';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook } = await import('@testing-library/react');

// Every family renders: the probe measures a font string by its length.
const probe = {
  font: '',
  textBaseline: '',
  measureText(this: { font: string }) {
    return { width: this.font.length };
  },
};
HTMLCanvasElement.prototype.getContext = (() => probe) as never;

const ROOT = resolve(import.meta.dir, '../../../../../..');
const DOCX = new Uint8Array(
  readFileSync(resolve(ROOT, 'crates/docx-edit/tests/fixtures/structured-export/principal.docx'))
);

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(ROOT, 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  )
);

afterAll(async () => {
  cleanup();
  if (ownsDom) await GlobalRegistrator.unregister();
});

async function pickerFonts(
  hosts: { preview: YrsDocxHost; full: YrsDocxHost },
  fontsFromFullDocument: boolean
): Promise<FontOption[][]> {
  const offered: FontOption[][] = [];
  const fontScope = createFontLoadScope();
  fontScope.loadFontsWithMapping = async () => {};
  const { result, unmount } = renderHook(() =>
    useDocumentLoader({
      documentBuffer: null,
      initialDocument: null,
      externalContent: false,
      history: useHistory<Document | null>(null),
      pagedEditorRef: { current: null },
      setLoadingState: () => {},
      setComments: () => {},
      setShowCommentsSidebar: () => {},
      onError: undefined,
      resetForNewDocument: () => {},
      commentsLoadedRef: { current: false },
      commentIdAllocator: createCommentIdAllocator(),
      setDocumentFonts: (fonts) => offered.push(fonts),
      fontScope,
      fontsFromFullDocument,
    })
  );
  await act(async () => {
    void result.current.loadBuffer(DOCX.slice().buffer);
  });
  const generation = result.current.yrsSeedGeneration;
  await act(async () => {
    result.current.acceptHostDocument(hosts.preview, generation, undefined, { preview: true });
  });
  await act(async () => {
    result.current.acceptHostDocument(hosts.full, generation);
  });
  unmount();
  fontScope.dispose();
  return offered;
}

test("the picker's document fonts come from the full document when a preview offers none", async () => {
  const previewSession = await createYrsSession();
  const fullSession = await createYrsSession();
  const hosts = {
    preview: previewSession.openDocxPreview(DOCX, 1)!,
    full: fullSession.openDocx(DOCX, true),
  };

  const [previewFonts, fullFonts] = await pickerFonts(hosts, false);
  expect(previewFonts!.length).toBeGreaterThan(0);
  expect(fullFonts!.length).toBeGreaterThan(0);

  expect(await pickerFonts(hosts, true)).toEqual([[], fullFonts!]);
  previewSession.destroy();
  fullSession.destroy();
});
