import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { Layout } from '@betteroffice/docx/layout/pagination';
import { getLayoutKernelInputs } from '@betteroffice/docx/editor';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { createYrsSession, type YrsSession } from '@betteroffice/docx/yrs';
import { PagedEditor, type PagedEditorRef } from './PagedEditor';
import type { YrsCoreSession } from './hooks/useYrsCoreSession';

const { act, cleanup, render, waitFor } = await import('@testing-library/react');

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);

let session: YrsSession;
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
  session = await createYrsSession({ clientId: 4242 });
  session.createStory('body', 'AXZ');
  session.applyRawOps('body', [{ op: 'format', index: 1, len: 1, attrs: { hidden: true } }]);
});

afterEach(() => {
  cleanup();
});

afterAll(async () => {
  session?.destroy();
  if (ownsDom) await GlobalRegistrator.unregister();
});

function yrsCore(): YrsCoreSession {
  return {
    session,
    storyBlocks: () => null,
    bodyBlocks: () => null,
    inputPositionMap: () => null,
    displayPositionToLoc: () => null,
    locToDisplayPosition: () => null,
    documentFromYrs: () => null,
    publishDirectInput: () => {},
  };
}

function measuredText(layout: Layout): string {
  const kernel = getLayoutKernelInputs(layout);
  return JSON.stringify(kernel?.measured ?? null);
}

test('showHiddenText reveals vanished runs in the paged layout', async () => {
  const layouts: Layout[] = [];
  const errors: Error[] = [];
  let ref: PagedEditorRef | null = null;
  const view = render(
    <PagedEditor
      ref={(instance) => {
        ref = instance;
      }}
      document={null}
      yrsCore={yrsCore()}
      measurementFontProvider={{ resolve: () => () => Promise.resolve(fontBytes) }}
      onError={(error) => {
        errors.push(error);
      }}
      onLayoutComputed={(layout) => {
        if (layout) layouts.push(layout);
      }}
    />
  );
  try {
    await waitFor(() => expect(layouts.length).toBeGreaterThan(0));
    expect(measuredText(layouts.at(-1)!)).not.toContain('"AXZ"');

    view.rerender(
      <PagedEditor
        ref={(instance) => {
          ref = instance;
        }}
        document={null}
        yrsCore={yrsCore()}
        measurementFontProvider={{ resolve: () => () => Promise.resolve(fontBytes) }}
        showHiddenText
        onError={(error) => {
          errors.push(error);
        }}
        onLayoutComputed={(layout) => {
          if (layout) layouts.push(layout);
        }}
      />
    );
    act(() => {
      ref!.relayout();
    });

    await waitFor(() => expect(measuredText(layouts.at(-1)!)).toContain('"AXZ"'));
  } finally {
    if (layouts.length === 0) console.log('layout errors:', errors.map(String));
    view.unmount();
  }
});
