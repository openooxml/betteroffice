import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { Layout } from '@betteroffice/docx/layout/pagination';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import {
  createYrsSession,
  proposalRevisionPreview,
  type DocxProposalResult,
  type YrsSession,
} from '@betteroffice/docx/yrs';
import { PagedEditor, type PagedEditorRef } from './PagedEditor';
import type { YrsCoreSession } from './hooks/useYrsCoreSession';
import { useRevisionPreview } from './hooks/useRevisionPreview';

const { act, cleanup, render, renderHook } = await import('@testing-library/react');

const WASM = resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const SUGGEST = { author: 'Atira', date: '2026-09-29T00:00:00Z' };

let fontBytes: ArrayBuffer;
const sessions: YrsSession[] = [];

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
  for (const session of sessions.splice(0)) session.destroy();
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function yrsCore(session: YrsSession): YrsCoreSession {
  return {
    session,
    sessionGeneration: 0,
    storyBlocks: () => null,
    bodyBlocks: () => null,
    inputPositionMap: () => null,
    displayPositionToLoc: () => null,
    locToDisplayPosition: () => null,
    documentFromYrs: () => null,
    publishDirectInput: () => {},
    scheduleCompatibilityWarm: () => {},
    cancelCompatibilityWarm: () => {},
  };
}

async function settleUntil(done: () => boolean): Promise<void> {
  const deadline = Date.now() + 2000;
  while (!done() && Date.now() < deadline) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
  }
  expect(done()).toBe(true);
}

function ok(result: DocxProposalResult) {
  if (!result.ok) throw new Error(`${result.failure.code}: ${result.failure.message}`);
  return result.snapshot;
}

test('proposal decisions reach the render env and lay the document out again', async () => {
  const session = await createYrsSession({ clientId: 4343 });
  sessions.push(session);
  const { paraId } = session.createStory('body', 'Hello world and more');
  const paragraph = {
    kind: 'session',
    sessionId: session.paragraphIdentities().sessionId,
    story: 'body',
    paraId,
  } as const;
  const layouts: Layout[] = [];
  const editor = createRef<PagedEditorRef>();
  render(
    <PagedEditor
      ref={editor}
      document={null}
      yrsCore={yrsCore(session)}
      measurementFontProvider={{ resolve: () => () => Promise.resolve(fontBytes) }}
      onLayoutComputed={(layout) => {
        if (layout) layouts.push(layout);
      }}
    />
  );
  const env = () => {
    const request = editor.current?.getLayoutRequest();
    return request
      ? (JSON.parse(request) as { renderEnv: Record<string, unknown> }).renderEnv
      : null;
  };
  await settleUntil(() => layouts.length > 0 && env() !== null);
  expect(env()).not.toHaveProperty('revisionPreview');

  await act(async () => {
    ok(
      session.proposeChanges({
        expectVersion: session.version(),
        proposals: [
          {
            id: 'a',
            paragraph,
            suggest: SUGGEST,
            op: 'replaceText',
            search: 'world',
            replaceWith: 'earth',
          },
          {
            id: 'b',
            paragraph,
            suggest: SUGGEST,
            op: 'replaceText',
            search: 'more',
            replaceWith: '',
          },
        ],
      })
    );
  });
  expect(env()).not.toHaveProperty('revisionPreview');

  const laidOut = layouts.length;
  let decided!: ReturnType<typeof ok>;
  await act(async () => {
    decided = ok(
      session.setProposalStates({
        expectVersion: session.version(),
        expectPreviewVersion: 0,
        changes: [
          { id: 'a', state: 'accepted' },
          { id: 'b', state: 'rejected' },
        ],
      })
    );
  });
  await settleUntil(() => layouts.length > laidOut);
  expect(env()?.revisionPreview).toEqual(proposalRevisionPreview(decided));

  await act(async () => {
    ok(
      session.setProposalStates({
        expectVersion: session.version(),
        expectPreviewVersion: 1,
        changes: [
          { id: 'a', state: 'proposed' },
          { id: 'b', state: 'proposed' },
        ],
      })
    );
  });
  await settleUntil(() => env() !== null && !('revisionPreview' in env()!));
});

test('the render env preview changes identity only when a decision does', async () => {
  const first = await createYrsSession({ clientId: 4344 });
  const second = await createYrsSession({ clientId: 4345 });
  sessions.push(first, second);
  const { paraId } = first.createStory('body', 'Hello world');
  const paragraph = {
    kind: 'session',
    sessionId: first.paragraphIdentities().sessionId,
    story: 'body',
    paraId,
  } as const;
  second.createStory('body', 'Other');
  const hook = renderHook(({ session }) => useRevisionPreview(session), {
    initialProps: { session: first as YrsSession | null },
  });
  const initial = hook.result.current;
  expect(initial).toEqual({ previewVersion: 0, revisionPreview: undefined });

  act(() => {
    ok(
      first.proposeChanges({
        expectVersion: first.version(),
        proposals: [
          { id: 'a', paragraph, suggest: SUGGEST, op: 'insertText', at: 'end', text: '!' },
        ],
      })
    );
  });
  expect(hook.result.current).toBe(initial);

  let decided!: ReturnType<typeof ok>;
  act(() => {
    decided = ok(
      first.setProposalStates({
        expectVersion: first.version(),
        expectPreviewVersion: 0,
        changes: [{ id: 'a', state: 'rejected' }],
      })
    );
  });
  const rejected = hook.result.current;
  expect(rejected).toEqual({
    previewVersion: 1,
    revisionPreview: proposalRevisionPreview(decided),
  });
  hook.rerender({ session: first });
  expect(hook.result.current).toBe(rejected);

  hook.rerender({ session: second });
  expect(hook.result.current).toEqual({ previewVersion: 0, revisionPreview: undefined });
  hook.rerender({ session: null });
  expect(hook.result.current.revisionPreview).toBeUndefined();
});
