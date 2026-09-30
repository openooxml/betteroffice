import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { useEffect } from 'react';
import type { Layout } from '@betteroffice/docx/layout/pagination';
import { createDisplayListQueries, type DisplayList } from '@betteroffice/docx/layout/render';
import {
  createCanvasHostProjector,
  createRenderedDomContext,
} from '@betteroffice/docx/plugin-api/RenderedDomContext';
import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import type { YrsLoc } from '@betteroffice/docx/yrs';
import { createDocxCommandController } from '../commands/createDocxCommandStore';
import { EngineWorker, lazyFixture } from '../components/DocxEditor/hooks/__fixtures__/lazyPages';
import { useRustDisplayList } from '../components/DocxEditor/hooks/useDisplayList';
import {
  markPresented,
  stampRevisionPreviewKey,
  stampSourceVersion,
} from '../components/DocxEditor/internals/layoutProvenance';
import type { PagedEditorRef } from '../components/DocxEditor/PagedEditor';
import type { SelectionState } from '../components/DocxEditor/types';
import { defineDocxPlugin } from './defineDocxPlugin';
import { createPluginGeometry } from './geometry';
import { renderedPreviewKey, type DocxProposalSnapshot } from './proposalPreview';
import type { DocxAnchorGeometryResult, DocxGeometryTarget } from './types';
import { useDocxPluginHost } from './useDocxPluginHost';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, renderHook, waitFor } = await import('@testing-library/react');
const originalWorker = globalThis.Worker;
const originalIdle = globalThis.requestIdleCallback;
const originalCancelIdle = globalThis.cancelIdleCallback;

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(
      readFileSync(
        resolve(import.meta.dir, '../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm')
      )
    )
  )
);

afterEach(() => {
  cleanup();
  globalThis.Worker = originalWorker;
  globalThis.requestIdleCallback = originalIdle;
  globalThis.cancelIdleCallback = originalCancelIdle;
});

afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

test('a visible proposal waits for its page build and paint, then notifies plugins with exact geometry', () =>
  visibleProposal(false)
);

test('scrolling before an edit reply builds newly visible placeholders and waits for paint', () =>
  visibleProposal(true)
);

async function visibleProposal(editDuringScroll: boolean): Promise<void> {
  let nextIdle = 1;
  globalThis.requestIdleCallback = (() => nextIdle++) as typeof requestIdleCallback;
  globalThis.cancelIdleCallback = () => {};
  const { engine, inputs, host } = lazyFixture(editDuringScroll ? undefined : 120);
  const pages = document.createElement('div');
  const layer = document.createElement('div');
  document.body.append(pages, layer);
  try {
    const paragraphs = JSON.parse(engine.paragraphs('body')) as { paraId: string; text: string }[];
    const paragraph = paragraphs.at(-1)!;
    const revisionOffset = editDuringScroll ? paragraph.text.length - 10 : 1;
    const measured = inputs.measured.find(
      ({ block }: { block: { id: string } }) => block.id === paragraph.paraId
    );
    const start = measured.block.pmStart + 1;
    const anchor = {
      kind: 'session' as const,
      sessionId: 'session',
      story: 'body',
      paraId: paragraph.paraId,
    };
    const snapshot: DocxProposalSnapshot = {
      version: 'v1',
      previewVersion: 0,
      proposals: [
        {
          id: 'proposal',
          state: 'accepted',
          paragraph: anchor,
          revisionIds: ['r1'],
          changed: true,
        },
      ],
    };
    Object.assign(host, {
      version: () => snapshot.version,
      getProposals: () => snapshot,
      resolveParagraphAnchor: () => ({ status: 'found', anchor }),
      paragraphSpans: () => [{ paraId: paragraph.paraId, length: paragraph.text.length }],
      listRevisions: () => [
        {
          revisionId: 'r1',
          kind: 'insertion',
          story: 'body',
          range: {
            start: { paraId: paragraph.paraId, offset: revisionOffset },
            end: { paraId: paragraph.paraId, offset: revisionOffset + 4 },
          },
        },
      ],
    });
    const editor = {
      getYrsSession: () => host,
      getSelectionRange: () => null,
      hasPendingInput: () => false,
      yrsLocToDisplayPosition: ({ offset }: YrsLoc) => start + offset,
    } as unknown as PagedEditorRef;
    const target: DocxGeometryTarget = { kind: 'proposal', id: 'proposal' };
    const events: (DocxAnchorGeometryResult | null)[] = [];
    const plugin = defineDocxPlugin({
      id: 'test.lazy-proposal',
      createState: () => null,
      onEvent(context, event) {
        if (event.type === 'layout-change') {
          events.push(context.geometry?.getAnchorGeometry(target) ?? null);
        }
      },
    });
    const options = {
      plugins: [plugin],
      pagedEditorRef: { current: editor },
      writeModeRef: { current: 'viewing' as const },
      mode: 'viewing' as const,
      readOnly: true,
      commands: createDocxCommandController(),
      session: host,
      loadGeneration: 0,
      layoutError: null,
      zoom: 1,
      canvasHostRef: { current: pages },
      overlayTarget: layer,
      selectionChangeSubscribersRef: { current: new Set<(state: SelectionState | null) => void>() },
      i18n: undefined,
      onRenderedDomContextReady: undefined,
    };
    stampSourceVersion(inputs.layout, snapshot.version);
    let autoPresent = true;
    let present: (() => void) | null = null;
    const overrides = { getInputs: () => inputs };
    const { result, unmount } = renderHook(() => {
      const display = useRustDisplayList(
        inputs.layout as Layout,
        overrides,
        undefined,
        undefined,
        host
      );
      const binding = useDocxPluginHost({ ...options, queries: display.queries });
      useEffect(() => {
        const queries = display.queries;
        if (!queries || !display.displayList) return;
        let active = true;
        void queries.whenReady().then(() => {
          if (!active) return;
          if (pages.children.length === 0) {
            pages.getBoundingClientRect = () => new DOMRect(0, 0, 500, 1000);
            layer.getBoundingClientRect = () => new DOMRect(0, 0, 500, 1000);
            for (const page of display.displayList!.pages) {
              const canvas = document.createElement('canvas');
              canvas.dataset.pageIndex = String(page.pageIndex);
              canvas.getBoundingClientRect = () =>
                new DOMRect(0, page.pageIndex * (page.height + 24), page.width, page.height);
              pages.append(canvas);
            }
          }
          // As in the editor, the context arrives before the frame paints.
          binding.onRenderedDomContext(
            createRenderedDomContext(pages, 1, {
              displayListQueries: queries,
              projector: createCanvasHostProjector(pages, queries, 1),
            }),
            queries
          );
          const paint = () => markPresented(pages, display.displayList!);
          if (autoPresent) {
            requestAnimationFrame(() => {
              if (active) paint();
            });
          } else {
            present = paint;
          }
        });
        return () => {
          active = false;
        };
      }, [display.queries, binding.onRenderedDomContext]);
      return { display, binding };
    });
    await act(async () => result.current.binding.overlayLayerRef(layer));
    const geometry = () => result.current.binding.activations[0]?.context.geometry;
    await waitFor(() => expect(geometry()?.getAnchorGeometry(target)).toMatchObject({ ok: true }));
    const last = result.current.display.frame!.pages.length - 1;
    expect(last).toBeGreaterThan(5);
    expect(result.current.display.frame!.displayList.pages[last]!.unbuilt).toBe(true);
    // A paragraph spanning pages falls back to its built first page.
    expect(geometry()!.getAnchorGeometry(target)).toMatchObject(
      editDuringScroll
        ? { ok: true, rects: [], anchor: { pageIndex: 0 } }
        : { ok: true, rects: [], anchor: { pageIndex: last, width: 0, height: 0 } }
    );

    const eventsBeforeWindow = events.length;
    const worker = EngineWorker.last!;
    let pendingEdit: ReturnType<typeof result.current.display.applyInput> | undefined;
    if (editDuringScroll) {
      await act(async () => {
        await result.current.display.settledDisplayList(null);
      });
      await waitFor(() =>
        expect(geometry()!.getAnchorGeometry(target)).toMatchObject({
          ok: true,
          rects: expect.arrayContaining([expect.objectContaining({ pageIndex: last })]),
        })
      );
      expect(result.current.display.frame!.displayList.pages.every((page) => !page.unbuilt)).toBe(
        true
      );
      engine.set_selection('body', paragraph.paraId, 1, paragraph.paraId, 1);
      worker.holdInputReplies = true;
      await act(async () => {
        pendingEdit = result.current.display.applyInput('New ');
      });
      await waitFor(() => expect(worker.heldInputReplies).toHaveLength(1));
      expect(worker.posted.at(-1)).toMatchObject({
        type: 'applyInput',
        displayWindow: [0, 5],
      });
    }
    worker.holdPageBuilds = true;
    await act(async () => result.current.display.setDisplayWindow(last, last + 1));
    if (editDuringScroll) {
      expect(worker.heldInputReplies).toHaveLength(1);
      expect(worker.heldPageBuilds).toEqual([]);
      worker.holdInputReplies = false;
      await act(async () => {
        worker.releaseInputReplies();
        expect(await pendingEdit!).not.toBeNull();
      });
      expect(result.current.display.error).toBeNull();
      expect(result.current.display.frame!.displayList.pages[last]!.unbuilt).toBe(true);
      expect(result.current.display.frame!.displayList.pages[0]!.unbuilt).toBeFalsy();
      expect(result.current.display.caret?.caretRect?.pageIndex).toBe(0);
    }
    expect(geometry()!.getAnchorGeometry(target)).toMatchObject({
      ok: false,
      failure: { code: 'layout-unavailable' },
    });
    await waitFor(() => expect(worker.heldPageBuilds).toHaveLength(1));
    await waitFor(() =>
      expect(events.slice(eventsBeforeWindow)).toContainEqual(
        expect.objectContaining({
          ok: false,
          failure: expect.objectContaining({ code: 'layout-unavailable' }),
        })
      )
    );
    expect(worker.posted.at(-1)).toMatchObject({ type: 'buildPages', pages: [last] });

    autoPresent = false;
    const eventsBeforeBuild = events.length;
    await act(async () => worker.releasePageBuilds());
    await waitFor(() =>
      expect(result.current.display.frame!.displayList.pages[last]!.unbuilt).toBeFalsy()
    );
    await waitFor(() => expect(present).not.toBeNull());
    expect(geometry()?.getAnchorGeometry(target)).toMatchObject({
      ok: false,
      failure: { code: 'layout-unavailable' },
    });
    await act(async () => present!());
    await waitFor(() =>
      expect(
        events.slice(eventsBeforeBuild).some((event) => event?.ok && event.rects.length > 0)
      ).toBe(true)
    );
    const exact = geometry()!.getAnchorGeometry(target);
    expect(exact).toMatchObject({ ok: true, anchor: { pageIndex: last } });
    if (!exact.ok) throw new Error(exact.failure.message);
    expect(exact.rects.length).toBeGreaterThan(0);

    if (editDuringScroll) {
      autoPresent = true;
      worker.holdPageBuilds = false;
      await act(async () => {
        expect(await result.current.display.applyInput('Next ')).not.toBeNull();
      });
      expect(worker.posted.filter((request) => request.type === 'applyInput').at(-1)).toMatchObject({
        displayWindow: [last, last + 1],
      });
      expect(result.current.display.error).toBeNull();
      expect(result.current.display.caret?.caretRect?.pageIndex).toBe(0);
      unmount();
      return;
    }

    const full = JSON.parse(engine.build_display_list_json(JSON.stringify(inputs))) as DisplayList;
    const fullQueries = createDisplayListQueries(full);
    await fullQueries.whenReady();
    stampRevisionPreviewKey(fullQueries, renderedPreviewKey(result.current.display.queries!));
    const reference = createPluginGeometry(
      geometry()!.layout,
      createRenderedDomContext(pages, 1, {
        displayListQueries: fullQueries,
        projector: createCanvasHostProjector(pages, fullQueries, 1),
      }),
      layer,
      () => true,
      () => null,
      fullQueries,
      () => ({ session: host, editor, presented: true })
    ).getAnchorGeometry(target);
    expect(reference.ok).toBe(true);
    if (!reference.ok) throw new Error(reference.failure.message);
    expect(exact.rects).toEqual(reference.rects);
    expect(exact.anchor).toEqual(reference.anchor);
    fullQueries.dispose();
    unmount();
  } finally {
    pages.remove();
    layer.remove();
    engine.free();
  }
}
