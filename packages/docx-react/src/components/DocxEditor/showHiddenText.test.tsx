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
import { PagedEditor } from './PagedEditor';
import type { YrsCoreSession } from './hooks/useYrsCoreSession';

const { act, cleanup, render } = await import('@testing-library/react');

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

async function settleLayoutsUntil(done: () => boolean): Promise<void> {
  const deadline = Date.now() + 2000;
  while (!done()) {
    if (Date.now() >= deadline) break;
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
  }
  expect(done()).toBe(true);
}

test('showHiddenText reveals vanished runs in the paged layout', async () => {
  const layouts: Layout[] = [];
  const errors: Error[] = [];
  const view = render(
    <PagedEditor
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
    await settleLayoutsUntil(() => layouts.length > 0);
    expect(measuredText(layouts.at(-1)!)).not.toContain('"AXZ"');

    view.rerender(
      <PagedEditor
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

    await settleLayoutsUntil(
      () => layouts.length > 0 && measuredText(layouts.at(-1)!).includes('"AXZ"')
    );
  } finally {
    if (layouts.length === 0) console.log('layout errors:', errors.map(String));
    view.unmount();
  }
});
