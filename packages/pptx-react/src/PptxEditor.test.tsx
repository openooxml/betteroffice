/**
 * `fonts` identity drives the effect that disposes and reopens the presentation,
 * so a caller building the array inline must not lose the open deck — its edits,
 * history and collaboration replica — on every unrelated re-render.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { useState } from 'react';
import type { ComponentProps } from 'react';
import { initWasm } from '@betteroffice/pptx';
import * as pptx from '@betteroffice/pptx';
import type {
  DeckSnapshot,
  PptxEditRequest,
  PptxEditResult,
  PptxFontFace,
  PptxPresenceCursor,
  SlideDisplayList,
} from '@betteroffice/pptx';
import type { PptxEditorApi, PptxEditorProps, PptxWorkerViewerApi } from './PptxEditor';
import { paintSelection, PptxEditor, SelectionOverlay } from './PptxEditor';
import { EditorToolbar, PptxCommandProvider, ToolbarCommandButton } from './index';
import { isMacPlatform, matchesChord } from './commands/descriptors';
import * as presenceRendering from './presence-rendering';

const mod = () => (isMacPlatform() ? { metaKey: true } : { ctrlKey: true });

type Equal<A, B> = (<T>() => T extends A ? 1 : 2) extends
  (<T>() => T extends B ? 1 : 2) ? true : false;

function assertType<T extends true>(_value?: T): void {}

it('preserves editor props and contextual API types across overloads', () => {
  assertType<Equal<Parameters<typeof PptxEditor>[0], PptxEditorProps>>();
  assertType<Equal<PptxEditorProps, Parameters<typeof PptxEditor>[0]>>();
  assertType<Equal<ComponentProps<typeof PptxEditor>, PptxEditorProps>>();
  assertType<Equal<PptxEditorProps, ComponentProps<typeof PptxEditor>>>();
  const check = (props: Partial<Parameters<typeof PptxEditor>[0]> = {}) => {
    const opened: PptxEditorApi[] = [];
    const editor = <PptxEditor fonts={[]} onReady={(api) => {
      assertType<Equal<typeof api, PptxEditorApi>>();
      opened.push(api);
    }} {...props} />;
    const viewer = <PptxEditor fonts={[]} readOnly experimentalWorkerOpen onReady={(api) => {
      assertType<Equal<typeof api, PptxWorkerViewerApi>>();
    }} />;
    return { editor, viewer };
  };
  void check;
});

const root = resolve(import.meta.dir, '../../..');

// the registrator writes one process-wide global set, so only the file that
// installed it may tear it down.
const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor, within } = await import(
  '@testing-library/react'
);

let fixture: Uint8Array;
let fontBytes: Uint8Array;

beforeAll(async () => {
  const [wasm, pptx, font] = await Promise.all([
    readFile(resolve(root, 'packages/pptx/src/wasm/generated/pptx_wasm_bg.wasm')),
    readFile(resolve(root, 'apps/demo/public/betteroffice-demo.pptx')),
    readFile(resolve(root, 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')),
  ]);
  await initWasm(wasm);
  fixture = pptx;
  fontBytes = font;
});

afterEach(cleanup);
afterAll(async () => {
  if (ownsDom && GlobalRegistrator.isRegistered) await GlobalRegistrator.unregister();
});

/** Disabled controls with a stated reason stay focusable and are marked `aria-disabled`. */
function isDisabled(element: HTMLElement): boolean {
  return element.getAttribute('aria-disabled') === 'true' || (element as HTMLButtonElement).disabled;
}

describe('shortcut matching', () => {
  it('needs exactly the chord\'s modifiers', () => {
    const press = (init: KeyboardEventInit) => new KeyboardEvent('keydown', { key: 'b', ...init });
    expect(matchesChord('Mod+B', press({ ctrlKey: true }), false)).toBe(true);
    expect(matchesChord('Mod+B', press({ ctrlKey: true, metaKey: true }), false)).toBe(false);
    expect(matchesChord('Mod+B', press({ metaKey: true, ctrlKey: true }), true)).toBe(false);
    expect(matchesChord('Mod+B', press({ ctrlKey: true, altKey: true }), false)).toBe(false);
  });
});

describe('PptxEditor slide layout cache', () => {
  function getContext(this: HTMLCanvasElement): CanvasRenderingContext2D {
    return {
      canvas: this,
      fillStyle: '#000000',
      save() {}, restore() {}, setTransform() {}, drawImage() {}, fillRect() {},
    } as unknown as CanvasRenderingContext2D;
  }

  for (const scenario of [
    { name: 'instant layouts', layoutMs: 0, failAt: undefined },
    { name: 'layouts exceeding the budget', layoutMs: 13, failAt: undefined },
    { name: 'a layout failure', layoutMs: 0, failAt: 3 },
  ]) {
    it(`batches deferred thumbnails with ${scenario.name} in a 50-slide deck`, async () => {
      const fonts = [{ family: 'Liberation Sans', bytes: fontBytes }];
      const originalOpen = pptx.openPresentation;
      const peer = originalOpen(fixture, { clientId: 9440, fonts });
      const calls: number[] = [];
      const errors: Error[] = [];
      const failure = new Error('layout failed');
      let clock = 0;
      let finishPaint!: () => void;
      const firstPaint = new Promise<void>((resolve) => { finishPaint = resolve; });
      const paintedFrames = new Map<HTMLCanvasElement, SlideDisplayList>();
      const restorers: Array<() => void> = [];
      let api: PptxEditorApi | undefined;
      let view: ReturnType<typeof render> | undefined;
      try {
        let count = peer.snapshot().slides.length;
        while (count < 50) peer.insertSlide(count++);
        const seed = peer.encodeStateAsUpdate();
        const expected = peer.snapshot().slides.map((_slide, index) => peer.layoutSlide(index));
        const now = spyOn(performance, 'now').mockImplementation(() => clock);
        restorers.push(() => now.mockRestore());
        const open = spyOn(pptx, 'openPresentation').mockImplementation((bytes, options) => {
          const handle = originalOpen(bytes, options);
          const originalLayout = handle.layoutSlide.bind(handle);
          const layout = spyOn(handle, 'layoutSlide').mockImplementation((index) => {
            calls.push(index);
            clock += scenario.layoutMs;
            if (index === scenario.failAt) throw failure;
            return originalLayout(index);
          });
          restorers.push(() => layout.mockRestore());
          return handle;
        });
        restorers.push(() => open.mockRestore());
        const paint = spyOn(pptx, 'paintSlide').mockImplementation((ctx, frame, _dpr, scale) => {
          paintedFrames.set(ctx.canvas, frame);
          return scale === 1 ? firstPaint : Promise.resolve();
        });
        restorers.push(() => paint.mockRestore());
        const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
          getContext as unknown as HTMLCanvasElement['getContext']
        );
        restorers.push(() => context.mockRestore());
        const presence = spyOn(presenceRendering, 'groupPresenceBySlide');
        restorers.push(() => presence.mockRestore());
        view = render(<PptxEditor file={fixture} fonts={fonts}
          collaboration={{ clientId: 9441, initialUpdate: seed }}
          onError={(error) => errors.push(error)}
          onReady={(ready) => { api = ready; }} />);
        await waitFor(() => expect(api).toBeDefined());
        await waitFor(() => expect(paint).toHaveBeenCalled());
        expect(calls).toEqual([0]);
        expect(presence).toHaveBeenCalled();
        const presenceCalls = presence.mock.calls.length;
        const cache = (api!.handle as unknown as Record<symbol, {
          activate(slideId: string, key: string): boolean;
        }>)[Symbol.for('@betteroffice/pptx/slide-layout-cache')];
        const batches: number[][] = [];
        let batchStart = calls.length;
        const originalActivate = cache.activate.bind(cache);
        const activate = spyOn(cache, 'activate').mockImplementation((id, key) => {
          batches.push(calls.slice(batchStart));
          batchStart = calls.length;
          return originalActivate(id, key);
        });
        restorers.push(() => activate.mockRestore());
        expect(activate).not.toHaveBeenCalled();
        await act(async () => { finishPaint(); await firstPaint; });
        const completed = scenario.failAt ?? 50;
        const attempted = scenario.failAt === undefined ? 50 : completed + 1;
        await waitFor(() => {
          const canvases = view!.container.querySelectorAll<HTMLCanvasElement>('aside canvas');
          expect(canvases).toHaveLength(completed);
          expected.slice(0, completed).forEach((frame, index) => {
            expect(paintedFrames.get(canvases[index])).toEqual(frame);
            expect(paint.mock.calls.filter(([ctx]) => ctx.canvas === canvases[index])).toHaveLength(1);
          });
          expect(calls).toHaveLength(attempted);
          expect(errors).toEqual(scenario.failAt === undefined ? [] : [failure]);
        }, { timeout: 15_000 });
        if (scenario.failAt !== undefined) {
          await act(async () => {
            await new Promise((resolve) => setTimeout(resolve, 0));
            await new Promise((resolve) => setTimeout(resolve, 0));
          });
        }
        if (scenario.layoutMs === 0) {
          expect(batches).toEqual([Array.from({ length: attempted - 1 }, (_, index) => index + 1)]);
        } else {
          expect(batches).toEqual(Array.from({ length: 49 }, (_, index) => [index + 1]));
        }
        expect(activate).toHaveBeenCalledTimes(batches.length);
        const activeId = api!.handle.snapshot().slides[0].id;
        expect(activate.mock.calls.every(([id]) => id === activeId)).toBe(true);
        expect(activate.mock.results.every((result) =>
          result.type === 'return' && result.value === true)).toBe(true);
        expect(calls).toEqual(Array.from({ length: attempted }, (_, index) => index));
        expect(errors).toEqual(scenario.failAt === undefined ? [] : [failure]);
        expect(presence).toHaveBeenCalledTimes(presenceCalls);
      } finally {
        finishPaint();
        try {
          view?.unmount();
          peer.dispose();
        } finally {
          for (const restore of restorers.reverse()) restore();
        }
      }
    }, 60_000);
  }

  it('refreshes an edited slide thumbnail after undo and redo from another slide', async () => {
    const fonts = [{ family: 'Liberation Sans', bytes: fontBytes }];
    const paintedFrames = new Map<HTMLCanvasElement, SlideDisplayList>();
    const restorers: Array<() => void> = [];
    let api: PptxEditorApi | undefined;
    let view: ReturnType<typeof render> | undefined;
    try {
      const paint = spyOn(pptx, 'paintSlide').mockImplementation((ctx, frame) => {
        paintedFrames.set(ctx.canvas, frame);
        return Promise.resolve();
      });
      restorers.push(() => paint.mockRestore());
      const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
        getContext as unknown as HTMLCanvasElement['getContext']
      );
      restorers.push(() => context.mockRestore());
      view = render(<PptxEditor file={fixture} fonts={fonts}
        collaboration={{ clientId: 9450 }}
        onReady={(ready) => { api = ready; }} />);
      await waitFor(() => expect(api).toBeDefined());
      const thumbnail = () => view!.container.querySelector<HTMLCanvasElement>('aside button:first-child canvas');
      await waitFor(() => expect(paintedFrames.get(thumbnail()!)).toBeDefined());
      const original = paintedFrames.get(thumbnail()!)!;
      const slide = api!.handle.snapshot().slides[0];
      const shape = slide.shapes.find((shape) => shape.textStories.length)!;
      const story = shape.textStories[0];
      const originalStory = api!.handle.story(story.id);
      await act(async () => {
        expect(api!.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 0 })).toBe(true);
        fireEvent.keyDown(view!.getByRole('application'), { key: 'X' });
        await api!.flushPendingInput();
      });
      await waitFor(() => {
        expect(paintedFrames.get(thumbnail()!)).toBeDefined();
        expect(paintedFrames.get(thumbnail()!)).not.toEqual(original);
      });
      const editedStory = api!.handle.story(story.id);
      expect(editedStory).not.toEqual(originalStory);
      await act(async () => { expect(api!.goToSlide(2)).toBe(true); });
      for (const direction of ['undo', 'redo'] as const) {
        await act(async () => {
          expect(await api!.commands.execute(direction, null)).toEqual({ ok: true, status: 'executed' });
        });
        expect(view!.container.querySelectorAll('aside button')[1].getAttribute('aria-current')).toBe('page');
        expect(api!.handle.story(story.id)).toEqual(direction === 'undo' ? originalStory : editedStory);
        const fresh = pptx.openPresentation(fixture, {
          clientId: 9451, fonts, initialUpdate: api!.handle.encodeStateAsUpdate(),
        });
        try {
          const expected = fresh.layoutSlide(0);
          await waitFor(() => expect(paintedFrames.get(thumbnail()!)).toEqual(expected));
        } finally {
          fresh.dispose();
        }
      }
    } finally {
      try {
        view?.unmount();
      } finally {
        for (const restore of restorers.reverse()) restore();
      }
    }
  }, 30_000);

  for (const scenario of ['ready navigation', 'undo restoration'] as const) {
    it(`completes uncached thumbnails after ${scenario}`, async () => {
      const fonts = [{ family: 'Liberation Sans', bytes: fontBytes }];
      const peer = pptx.openPresentation(fixture, { clientId: 9420, fonts });
      let count = peer.snapshot().slides.length;
      while (count < 12) peer.insertSlide(count++);
      const seed = peer.encodeStateAsUpdate();
      const paintedFrames = new Map<HTMLCanvasElement, SlideDisplayList>();
      let finishPaint!: () => void;
      const firstPaint = new Promise<void>((resolve) => { finishPaint = resolve; });
      const paint = spyOn(pptx, 'paintSlide').mockImplementation((ctx, frame, _dpr, scale) => {
        paintedFrames.set(ctx.canvas, frame);
        return scale === 1 ? firstPaint : Promise.resolve();
      });
      const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
        getContext as unknown as HTMLCanvasElement['getContext']
      );
      let api: PptxEditorApi | undefined;
      let view: ReturnType<typeof render> | undefined;
      try {
        view = render(<PptxEditor file={fixture} fonts={fonts}
          collaboration={{ clientId: 9421, initialUpdate: seed }}
          onReady={(ready) => {
            api = ready;
            if (scenario === 'ready navigation') expect(ready.goToSlide(2)).toBe(true);
          }} />);
        await waitFor(() => expect(api).toBeDefined());
        await waitFor(() => expect(paint).toHaveBeenCalled());
        if (scenario === 'undo restoration') {
          await act(async () => {
            api!.handle.deleteSlide(api!.handle.snapshot().slides[11].id);
            api!.refresh();
            expect(await api!.commands.execute('undo', null)).toEqual({ ok: true, status: 'executed' });
          });
          expect(api!.handle.snapshot().slides).toHaveLength(12);
        }
        await act(async () => { finishPaint(); await firstPaint; });
        const fresh = pptx.openPresentation(fixture, {
          clientId: 9422, fonts, initialUpdate: api!.handle.encodeStateAsUpdate(),
        });
        try {
          const expected = fresh.snapshot().slides.map((_slide, index) => fresh.layoutSlide(index));
          await waitFor(() => {
            const canvases = view!.container.querySelectorAll<HTMLCanvasElement>('aside canvas');
            expect(canvases).toHaveLength(expected.length);
            expected.forEach((frame, index) => expect(paintedFrames.get(canvases[index])).toEqual(frame));
          }, { timeout: 15_000 });
        } finally {
          fresh.dispose();
        }
      } finally {
        finishPaint();
        view?.unmount();
        peer.dispose();
        paint.mockRestore();
        context.mockRestore();
      }
    }, 30_000);
  }

  for (const action of ['typing', 'navigation'] as const) {
    it(`completes every thumbnail after ${action} supersedes a running pass`, async () => {
      const fonts = [{ family: 'Liberation Sans', bytes: fontBytes }];
      const originalOpen = pptx.openPresentation;
      const peer = originalOpen(fixture, { clientId: 9430, fonts });
      let count = peer.snapshot().slides.length;
      while (count < 12) peer.insertSlide(count++);
      const seed = peer.encodeStateAsUpdate();
      const calls: number[] = [];
      let clock = 0;
      const now = spyOn(performance, 'now').mockImplementation(() => clock);
      let finishPaint!: () => void;
      const firstPaint = new Promise<void>((resolve) => { finishPaint = resolve; });
      let finishThumbnail!: () => void;
      const firstThumbnail = new Promise<void>((resolve) => { finishThumbnail = resolve; });
      const open = spyOn(pptx, 'openPresentation').mockImplementation((bytes, options) => {
        const handle = originalOpen(bytes, options);
        const layout = handle.layoutSlide.bind(handle);
        handle.layoutSlide = (index) => {
          const frame = layout(index);
          calls.push(index);
          clock += 13;
          if (index === 1) finishThumbnail();
          return frame;
        };
        return handle;
      });
      const paintedFrames = new Map<HTMLCanvasElement, SlideDisplayList>();
      const paint = spyOn(pptx, 'paintSlide').mockImplementation((ctx, frame, _dpr, scale) => {
        paintedFrames.set(ctx.canvas, frame);
        return scale === 1 ? firstPaint : Promise.resolve();
      });
      const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
        getContext as unknown as HTMLCanvasElement['getContext']
      );
      let api: PptxEditorApi | undefined;
      let view: ReturnType<typeof render> | undefined;
      try {
        view = render(<PptxEditor file={fixture} fonts={fonts}
          collaboration={{ clientId: 9431, initialUpdate: seed }}
          onReady={(ready) => { api = ready; }} />);
        await waitFor(() => expect(api).toBeDefined());
        await waitFor(() => expect(paint).toHaveBeenCalled());
        expect(calls).toEqual([0]);
        await act(async () => {
          finishPaint();
          await firstThumbnail;
          expect(calls).toEqual([0, 1]);
          if (action === 'typing') {
            const slide = api!.handle.snapshot().slides[0];
            const shape = slide.shapes.find((shape) => shape.textStories.length)!;
            const story = shape.textStories[0];
            expect(api!.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 0 })).toBe(true);
            fireEvent.keyDown(view!.getByRole('application'), { key: 'X' });
            await api!.flushPendingInput();
            expect(api!.handle.story(story.id).paragraphs[0].runs[0].text.startsWith('X')).toBe(true);
          } else {
            expect(api!.goToSlide(12)).toBe(true);
            expect(calls[calls.length - 1]).toBe(11);
          }
        });
        const fresh = originalOpen(fixture, {
          clientId: 9432, fonts, initialUpdate: api!.handle.encodeStateAsUpdate(),
        });
        try {
          const expected = fresh.snapshot().slides.map((_slide, index) => fresh.layoutSlide(index));
          await waitFor(() => {
            const canvases = view!.container.querySelectorAll<HTMLCanvasElement>('aside canvas');
            expect(canvases).toHaveLength(expected.length);
            expected.forEach((frame, index) => expect(paintedFrames.get(canvases[index])).toEqual(frame));
          }, { timeout: 15_000 });
        } finally {
          fresh.dispose();
        }
      } finally {
        finishPaint();
        view?.unmount();
        peer.dispose();
        open.mockRestore();
        paint.mockRestore();
        context.mockRestore();
        now.mockRestore();
      }
    }, 30_000);
  }

  it('computes one active key and at most one layout per typed character in a 50-slide deck', async () => {
    const fonts = [{ family: 'Liberation Sans', bytes: fontBytes }];
    const peer = pptx.openPresentation(fixture, { clientId: 9410, fonts });
    let count = peer.snapshot().slides.length;
    while (count < 50) peer.insertSlide(count++);
    const seed = peer.encodeStateAsUpdate();
    const paint = spyOn(pptx, 'paintSlide').mockResolvedValue(undefined);
    const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
      getContext as unknown as HTMLCanvasElement['getContext']
    );
    let api: PptxEditorApi | undefined;
    let view: ReturnType<typeof render> | undefined;
    try {
      view = render(<PptxEditor file={fixture} fonts={fonts}
        collaboration={{ clientId: 9411, initialUpdate: seed }}
        onReady={(ready) => { api = ready; }} />);
      await waitFor(() => expect(api).toBeDefined());
      await waitFor(() => expect(view!.container.querySelectorAll('aside canvas')).toHaveLength(50), {
        timeout: 15_000,
      });
      const slide = api!.handle.snapshot().slides[0];
      const shape = slide.shapes.find((shape) => shape.textStories.length)!;
      const story = shape.textStories[0];
      await act(async () => {
        expect(api!.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 0 })).toBe(true);
      });
      const cache = (api!.handle as unknown as Record<symbol, {
        snapshot(): unknown;
        key(index: number): string;
      }>)[Symbol.for('@betteroffice/pptx/slide-layout-cache')];
      const key = spyOn(cache, 'key');
      const snapshot = spyOn(cache, 'snapshot');
      const layout = spyOn(api!.handle, 'layoutSlide');
      try {
        for (const character of 'abcde') {
          key.mockClear();
          snapshot.mockClear();
          layout.mockClear();
          await act(async () => {
            fireEvent.keyDown(view!.getByRole('application'), { key: character });
            await api!.flushPendingInput();
            await new Promise((resolve) => setTimeout(resolve, 0));
          });
          expect(key.mock.calls.map((call) => call[0])).toEqual([0]);
          expect(snapshot).not.toHaveBeenCalled();
          expect(layout.mock.calls.length).toBeLessThanOrEqual(1);
          expect(layout.mock.calls.map((call) => call[0])).toEqual([0]);
          expect(view.container.querySelectorAll('aside canvas')).toHaveLength(50);
        }
        const text = api!.handle.story(story.id).paragraphs.flatMap((paragraph) =>
          paragraph.runs.map((run) => run.text)).join('');
        expect(text.startsWith('abcde')).toBe(true);
      } finally {
        layout.mockRestore();
        snapshot.mockRestore();
        key.mockRestore();
      }
    } finally {
      view?.unmount();
      peer.dispose();
      paint.mockRestore();
      context.mockRestore();
    }
  }, 60_000);

  it('paints the active slide first and lays out only damaged slides in a 50-slide deck', async () => {
    const fonts = [{ family: 'Liberation Sans', bytes: fontBytes }];
    const originalOpen = pptx.openPresentation;
    const peer = originalOpen(fixture, { clientId: 9400, fonts });
    let count = peer.snapshot().slides.length;
    while (count < 50) peer.insertSlide(count++);
    const seed = peer.encodeStateAsUpdate();
    const freshFrames = peer.snapshot().slides.map((_slide, index) => peer.layoutSlide(index));
    const calls: number[] = [];
    const frames: SlideDisplayList[] = [];
    let api: PptxEditorApi | undefined;
    let finishPaint!: () => void;
    const firstPaint = new Promise<void>((resolve) => { finishPaint = resolve; });
    let painted = false;
    const open = spyOn(pptx, 'openPresentation').mockImplementation((bytes, options) => {
      const handle = originalOpen(bytes, options);
      const layout = handle.layoutSlide.bind(handle);
      handle.layoutSlide = (index) => {
        calls.push(index);
        return layout(index);
      };
      return handle;
    });
    const paint = spyOn(pptx, 'paintSlide').mockImplementation((_ctx, frame, _dpr, scale) => {
      if (scale === 1) frames.push(frame);
      if (!painted) {
        painted = true;
        return firstPaint;
      }
      return Promise.resolve();
    });
    const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
      getContext as unknown as HTMLCanvasElement['getContext']
    );
    let view: ReturnType<typeof render> | undefined;
    try {
      view = render(<PptxEditor file={fixture} fonts={fonts}
        collaboration={{ clientId: 9401, initialUpdate: seed }}
        onReady={(ready) => { api = ready; }} />);
      await waitFor(() => expect(api).toBeDefined());
      await waitFor(() => expect(painted).toBe(true));
      expect(calls).toEqual([0]);
      expect(frames[0]).toEqual(freshFrames[0]);
      await act(async () => { finishPaint(); await firstPaint; });
      const thumbnails = () => view!.container.querySelectorAll('aside canvas');
      await waitFor(() => expect(thumbnails()).toHaveLength(50), { timeout: 15_000 });
      expect(calls).toEqual(Array.from({ length: 50 }, (_, index) => index));
      await waitFor(() => {
        const paintedFrames = paint.mock.calls.map((call) => call[1]);
        for (const frame of freshFrames) expect(paintedFrames).toContainEqual(frame);
      });

      const canvas = view.getByTestId('pptx-slide-canvas');
      const frame = freshFrames[0];
      canvas.getBoundingClientRect = () => new DOMRect(0, 0, frame.width, frame.height);
      canvas.setPointerCapture = () => {};
      calls.length = 0;
      let point: { x: number; y: number } | undefined;
      for (let y = 0; y < frame.height && !point; y += 16) {
        for (let x = 0; x < frame.width; x += 16) {
          if (api!.getPositionAtPoint(x, y)?.kind === 'text') { point = { x, y }; break; }
        }
      }
      expect(point).toBeDefined();
      expect(api!.getPositionAtPoint(point!.x, point!.y)).toEqual({
        ...api!.handle.hitTest(point!.x, point!.y)!, slide: 1, slideId: peer.snapshot().slides[0].id,
      });
      fireEvent.pointerDown(canvas, {
        isPrimary: true, button: 0, pointerId: 9, clientX: point!.x, clientY: point!.y,
      });
      fireEvent.pointerUp(canvas, { pointerId: 9, clientX: point!.x, clientY: point!.y });
      expect(calls).toHaveLength(0);

      const slide = peer.snapshot().slides[0];
      const shape = slide.shapes.find((shape) => shape.textStories.length)!;
      const story = shape.textStories[0];
      peer.insertText(story.id, 0, 'Remote ');
      await act(async () => {
        api!.handle.applyUpdate(peer.encodeDiff(api!.handle.encodeStateVector()));
      });
      await waitFor(() => expect(thumbnails()).toHaveLength(50));
      expect(calls).toEqual([0]);
      expect(frames[frames.length - 1]).toEqual(peer.layoutSlide(0));
      api!.getPositionAtPoint(point!.x, point!.y);
      expect(calls).toEqual([0]);

      calls.length = 0;
      const lastSlide = peer.snapshot().slides[49];
      peer.addTextBox(lastSlide.id, {
        name: 'Remote text', rect: { x: 100_000, y: 100_000, width: 1_000_000, height: 500_000 },
        text: 'Thumbnail change',
        style: {},
      });
      await act(async () => {
        api!.handle.applyUpdate(peer.encodeDiff(api!.handle.encodeStateVector()));
        expect(api!.goToSlide(1)).toBe(true);
      });
      await waitFor(() => expect(thumbnails()).toHaveLength(50));
      expect(calls).toEqual([49]);
      expect(frames[frames.length - 1]).toEqual(peer.layoutSlide(0));
      const changedThumbnail = peer.layoutSlide(49);
      await waitFor(() => expect(paint.mock.calls.map((call) => call[1])).toContainEqual(changedThumbnail));

      calls.length = 0;
      await act(async () => {
        const result = await api!.applyEdits({
          expectVersion: api!.handle.version(),
          steps: [{ op: 'insertText', at: 'start', text: 'Host ', target: {
            kind: 'range', slideId: slide.id, shapeId: shape.id, storyId: story.id, start: 0, end: 0,
          } }],
        });
        expect(result).toMatchObject({ ok: true, applied: true });
      });
      await waitFor(() => expect(thumbnails()).toHaveLength(50));
      expect(calls).toEqual([0]);
      const fresh = originalOpen(fixture, {
        clientId: 9402, fonts, initialUpdate: api!.handle.encodeStateAsUpdate(),
      });
      try { expect(frames[frames.length - 1]).toEqual(fresh.layoutSlide(0)); }
      finally { fresh.dispose(); }
      calls.length = 0;
      await act(async () => { api!.refresh(); });
      expect(calls).toHaveLength(0);
      await act(async () => { api!.goToSlide(25); });
      expect(calls).toEqual([24]);
      api!.getPositionAtPoint(10, 10);
      expect(calls).toEqual([24]);
    } finally {
      finishPaint();
      view?.unmount();
      peer.dispose();
      open.mockRestore();
      paint.mockRestore();
      context.mockRestore();
    }
  }, 60_000);
});

describe('PptxEditor PNG export', () => {
  const cases = [
    ['download', 'downloads the current slide'],
    ['unmount', 'discards an export after the deck is closed'],
    ['replace', 'discards an export after the deck is replaced'],
    ['stale-error', 'does not report an old export error to the replacement deck'],
    ['error', 'reports export errors while the deck is still open'],
  ] as const;
  for (const [scenario, title] of cases) {
    it(
      title,
      async () => {
        const originalCanvas = globalThis.OffscreenCanvas;
        const originalBitmap = globalThis.createImageBitmap;
        const downloads: string[] = [];
        const errors: Error[] = [];
        const clicked = spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (
          this: HTMLAnchorElement
        ) {
          downloads.push(this.download);
        });
        const createUrl = spyOn(URL, 'createObjectURL').mockReturnValue('blob:slide-png');
        const revokeUrl = spyOn(URL, 'revokeObjectURL').mockImplementation(() => {});
        let finishEncoding: ((blob: Blob) => void) | undefined;
        let failEncoding: ((error: Error) => void) | undefined;
        const encoded = new Promise<Blob>((resolve, reject) => {
          finishEncoding = resolve;
          failEncoding = reject;
        });
        let encoding = false;
        class ExportCanvas {
          getContext() {
            return new Proxy({} as Record<string, unknown>, {
              get(target, property) {
                if (property in target) return target[property as string];
                if (property === 'createLinearGradient' || property === 'createRadialGradient') {
                  return () => ({ addColorStop() {} });
                }
                return () => {};
              },
            });
          }
          convertToBlob() {
            encoding = true;
            return encoded;
          }
        }
        globalThis.OffscreenCanvas = ExportCanvas as unknown as typeof OffscreenCanvas;
        globalThis.createImageBitmap = (() =>
          Promise.resolve({} as ImageBitmap)) as typeof createImageBitmap;
        try {
          const opened: PptxEditorApi[] = [];
          const view = render(
            <PptxEditor
              file={fixture}
              fileName="report.pptx"
              fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]}
              clientId={9102}
              onReady={(api) => opened.push(api)}
              onError={(error) => errors.push(error)}
            />
          );
          await waitFor(() => expect(opened.length).toBe(1), { timeout: 15_000 });
          const exported = opened[0].commands.execute('exportPng', null);
          await waitFor(() => expect(encoding).toBe(true));
          if (scenario === 'unmount') view.unmount();
          if (scenario === 'replace' || scenario === 'stale-error') {
            view.rerender(
              <PptxEditor
                file={fixture.slice()}
                fileName="replacement.pptx"
                fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]}
                clientId={9102}
                onReady={(api) => opened.push(api)}
                onError={(error) => errors.push(error)}
              />
            );
            await waitFor(() => expect(opened.length).toBe(2), { timeout: 15_000 });
          }
          const encodingError = new Error('PNG encoding failed');
          await act(async () => {
            if (scenario === 'stale-error' || scenario === 'error') {
              failEncoding!(encodingError);
            } else {
              finishEncoding!(new Blob(['png'], { type: 'image/png' }));
            }
            await encoded.catch(() => {});
          });
          expect(downloads).toEqual(scenario === 'download' ? ['report-slide-1.png'] : []);
          expect(errors).toEqual(scenario === 'error' ? [encodingError] : []);
          expect(await exported).toMatchObject(
            scenario === 'download'
              ? { ok: true, status: 'executed' }
              : { ok: false, failure: { code: scenario === 'error' ? 'command-failed' : 'document-replaced' } }
          );
        } finally {
          cleanup();
          clicked.mockRestore();
          createUrl.mockRestore();
          revokeUrl.mockRestore();
          globalThis.OffscreenCanvas = originalCanvas;
          globalThis.createImageBitmap = originalBitmap;
        }
      },
      60_000
    );
  }
});

describe('PptxEditor insert image', () => {
  it(
    'adds a picture shape sized from the file and saves it back out',
    async () => {
      const originalImage = globalThis.Image;
      class FakeImage {
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        naturalWidth = 400;
        naturalHeight = 200;
        set src(_value: string) {
          queueMicrotask(() => this.onload?.());
        }
      }
      globalThis.Image = FakeImage as unknown as typeof Image;
      try {
        const opened: PptxEditorApi[] = [];
        const errors: Error[] = [];
        const view = render(
          <PptxEditor
            file={fixture}
            fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]}
            clientId={9103}
            onReady={(api) => opened.push(api)}
            onError={(error) => errors.push(error)}
          />
        );
        await waitFor(() => expect(opened.length).toBe(1), { timeout: 15_000 });
        const before = opened[0].handle.snapshot().slides[0].shapes.length;

        const bytes = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
        const file = new File([bytes], 'logo.png', { type: 'image/png' });
        const input = view.getByTestId('pptx-insert-image-input') as HTMLInputElement;
        await act(async () => {
          fireEvent.change(input, { target: { files: [file] } });
          await new Promise((resolve) => setTimeout(resolve, 0));
        });

        await waitFor(() => {
          expect(opened[0].handle.snapshot().slides[0].shapes.length).toBe(before + 1);
        });
        const shapes = opened[0].handle.snapshot().slides[0].shapes;
        const added = shapes[shapes.length - 1];
        expect(added.kind).toBe('picture');
        expect(added.name).toBe('logo.png');
        expect(Math.round(added.width / added.height)).toBe(2);
        expect(() => opened[0].handle.save()).not.toThrow();
        expect(errors).toEqual([]);
      } finally {
        cleanup();
        globalThis.Image = originalImage;
      }
    },
    30_000
  );

  it(
    'lands on the slide it was chosen for, not the one showing when decoding finishes',
    async () => {
      const originalImage = globalThis.Image;
      let finishDecoding: (() => void) | undefined;
      class PausedImage {
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        naturalWidth = 400;
        naturalHeight = 200;
        set src(value: string) {
          if (value.startsWith('data:')) finishDecoding = () => this.onload?.();
          else queueMicrotask(() => this.onload?.());
        }
      }
      globalThis.Image = PausedImage as unknown as typeof Image;
      try {
        const opened: PptxEditorApi[] = [];
        const view = render(
          <PptxEditor
            file={fixture}
            fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]}
            clientId={9104}
            onReady={(api) => opened.push(api)}
          />
        );
        await waitFor(() => expect(opened.length).toBe(1), { timeout: 15_000 });

        const bytes = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4]);
        const file = new File([bytes], 'logo.png', { type: 'image/png' });
        const input = view.getByTestId('pptx-insert-image-input') as HTMLInputElement;
        fireEvent.change(input, { target: { files: [file] } });

        await waitFor(() => expect(finishDecoding).toBeDefined());

        await act(async () => {
          expect(opened[0].goToSlide(2)).toBe(true);
        });

        await act(async () => {
          finishDecoding!();
          await Promise.resolve();
        });

        const hasLogo = (slideIndex: number) =>
          opened[0].handle
            .snapshot()
            .slides[slideIndex].shapes.some((shape) => shape.name === 'logo.png');
        await waitFor(() => expect(hasLogo(0)).toBe(true));
        expect(hasLogo(1)).toBe(false);
        expect(view.container.querySelector('button[aria-current="page"]')?.textContent).toContain('2');
      } finally {
        cleanup();
        globalThis.Image = originalImage;
      }
    },
    30_000
  );
});

describe('PptxEditor pending image insertion', () => {
  for (const scenario of ['replace', 'unmount', 'readOnly', 'decode-error', 'stale-error'] as const) {
    it(`handles ${scenario} while an image decodes`, async () => {
      const originalImage = globalThis.Image;
      let finish: (() => void) | undefined;
      class PausedImage {
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        naturalWidth = 400;
        naturalHeight = 200;
        set src(value: string) {
          if (value.startsWith('data:')) {
            finish = () => scenario.endsWith('error') ? this.onerror?.() : this.onload?.();
          } else queueMicrotask(() => this.onload?.());
        }
      }
      globalThis.Image = PausedImage as unknown as typeof Image;
      const opened: PptxEditorApi[] = [];
      const errors: Error[] = [];
      const props = {
        file: fixture,
        fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
        clientId: 9110,
        onReady: (api: PptxEditorApi) => opened.push(api),
        onError: (error: Error) => errors.push(error),
      };
      try {
        const view = render(<PptxEditor {...props} />);
        await waitFor(() => expect(opened).toHaveLength(1), { timeout: 15_000 });
        const before = opened[0].handle.snapshot();
        const file = new File([Uint8Array.from([0x89, 0x50, 0x4e, 0x47])], 'pending.png', { type: 'image/png' });
        fireEvent.change(view.getByTestId('pptx-insert-image-input'), { target: { files: [file] } });
        await waitFor(() => expect(finish).toBeDefined());
        if (scenario === 'replace' || scenario === 'stale-error') {
          view.rerender(<PptxEditor {...props} file={new Uint8Array(fixture)} />);
          await waitFor(() => expect(opened).toHaveLength(2), { timeout: 15_000 });
        } else if (scenario === 'unmount') {
          view.unmount();
        } else if (scenario === 'readOnly') {
          view.rerender(<PptxEditor {...props} readOnly />);
        }
        await act(async () => { finish!(); await Promise.resolve(); });
        if (scenario !== 'unmount') expect(opened[opened.length - 1].handle.snapshot()).toEqual(before);
        expect(errors).toHaveLength(scenario === 'decode-error' ? 1 : 0);
      } finally {
        cleanup();
        globalThis.Image = originalImage;
      }
    }, 30_000);
  }
});

describe('PptxEditor font stability', () => {
  // scanning a real font is what makes this slow; the budget is generous so a
  // loaded CI machine reports the assertion rather than a timeout.
  it(
    'keeps the open presentation across renders when fonts are rebuilt inline',
    async () => {
      const opened: PptxEditorApi[] = [];
      const onReady = (api: PptxEditorApi) => opened.push(api);
      const face = (): PptxFontFace[] => [
        { family: 'Liberation Sans', bytes: Uint8Array.from(fontBytes) },
      ];

      const { rerender } = render(
        <PptxEditor file={fixture} fonts={face()} clientId={9101} onReady={onReady} />
      );
      await waitFor(() => expect(opened.length).toBe(1), { timeout: 15_000 });

      await act(async () => {
        rerender(<PptxEditor file={fixture} fonts={face()} clientId={9101} onReady={onReady} />);
      });

      expect(opened.length).toBe(1);
      expect(() => opened[0].handle.snapshot()).not.toThrow();
    },
    60_000
  );
});

describe('PptxEditor host integration', () => {
  it('keeps viewing mode navigable without exposing user mutations', async () => {
    let api: PptxEditorApi | undefined;
    const cursors: Array<PptxPresenceCursor | null> = [];
    const presence = {
      peers: [],
      setCursor: (cursor: PptxPresenceCursor | null) => cursors.push(cursor),
      onPresence: () => () => {},
    };
    const view = render(
      <PptxEditor
        file={fixture}
        fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]}
        collaboration={{ clientId: 9120, presence }}
        initialSlide={2}
        onReady={(ready) => {
          api = ready;
        }}
        readOnly
      />
    );
    await act(async () => {
      await waitFor(() => expect(api).toBeDefined(), { timeout: 15_000 });
    });
    const snapshot = api!.handle.snapshot();
    expect(snapshot.slides.length).toBeGreaterThan(1);
    expect(view.queryByTestId('pptx-editor-toolbar')).toBeNull();
    expect((view.getByTestId('pptx-notes-textarea') as HTMLTextAreaElement).disabled).toBe(true);
    expect(view.container.querySelector('button[aria-current="page"]')?.textContent).toContain('2');

    await act(async () => {
      expect(api!.goToSlide(1)).toBe(true);
    });
    await waitFor(() =>
      expect(view.container.querySelector('button[aria-current="page"]')?.textContent).toContain('1')
    );
    const shape = snapshot.slides[0].shapes.find((candidate) => candidate.textStories.length > 0)!;
    const story = shape.textStories[0];
    const end = api!.handle.story(story.id).length;
    await act(async () => {
      expect(
        api!.selectText({
          slide: 1,
          shapeId: shape.id,
          storyId: story.id,
          start: 0,
          end,
        })
      ).toBe(true);
    });
    await waitFor(() => expect(cursors[cursors.length - 1]?.shapeId).toBe(shape.id));

    const stage = view.getByRole('application');
    fireEvent.keyDown(stage, { key: 'x' });
    fireEvent.keyDown(stage, { key: 'Backspace' });
    fireEvent.keyDown(stage, { key: 'z', ctrlKey: true });
    expect(api!.handle.snapshot()).toEqual(snapshot);

    await act(async () => api!.clearSelection());
    await waitFor(() => expect(cursors[cursors.length - 1]?.shapeId).toBeUndefined());
    expect(api!.goToSlide(0)).toBe(false);
    expect(
      api!.selectText({
        slide: 1,
        shapeId: shape.id,
        storyId: story.id,
        start: 0,
        end: Number.MAX_SAFE_INTEGER,
      })
    ).toBe(false);
  }, 60_000);
});

describe('PptxEditor viewing transitions', () => {
  it('keeps the document open and restores editing when viewing mode ends', async () => {
    const opened: PptxEditorApi[] = [];
    const onReady = (api: PptxEditorApi) => opened.push(api);
    const props = {
      file: fixture,
      fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
      onReady,
    };
    const view = render(
      <PptxEditor {...props} initialSlide={Number.MAX_SAFE_INTEGER} readOnly />
    );
    await act(async () => {
      await waitFor(() => expect(opened).toHaveLength(1), { timeout: 15_000 });
    });
    const api = opened[0];
    const before = api.handle.snapshot();
    const last = before.slides.length;
    expect(
      view.container.querySelector('button[aria-current="page"]')?.textContent
    ).toContain(String(last));
    await act(async () => {
      expect(api.goToSlide(Number.NaN)).toBe(false);
      expect(api.goToSlide(last + 1)).toBe(false);
    });
    expect(
      view.container.querySelector('button[aria-current="page"]')?.textContent
    ).toContain(String(last));
    await act(async () => {
      view.rerender(<PptxEditor {...props} initialSlide={1} />);
    });
    expect(opened).toHaveLength(1);
    expect(view.getByTestId('pptx-editor-toolbar')).toBeDefined();
    expect(
      view.container.querySelector('button[aria-current="page"]')?.textContent
    ).toContain(String(last));
    const shape = before.slides[0].shapes.find(
      (candidate) => candidate.textStories.length > 0
    )!;
    const story = shape.textStories[0];
    await act(async () => {
      expect(
        api.selectText({
          slide: 1,
          shapeId: shape.id,
          storyId: story.id,
          start: 0,
          end: 0,
        })
      ).toBe(true);
    });
    fireEvent.keyDown(view.getByRole('application'), { key: 'x' });
    expect(api.handle.story(story.id).paragraphs[0].runs[0].text).toStartWith('x');
    await act(async () => {
      view.rerender(<PptxEditor {...props} readOnly />);
    });
    const edited = api.handle.snapshot();
    fireEvent.keyDown(view.getByRole('application'), { key: 'Backspace' });
    expect(api.handle.snapshot()).toEqual(edited);
    expect(opened).toHaveLength(1);
  }, 60_000);
});

describe('PptxEditor caret painting', () => {
  const frame: SlideDisplayList = {
    contractVersion: 1,
    width: 320,
    height: 180,
    primitives: [
      {
        kind: 'textBox',
        objectId: 1,
        shapeId: 'shape',
        storyId: 'story',
        x: 20,
        y: 10,
        w: 200,
        h: 80,
        anchor: 'top',
        paragraphs: [],
        lines: [
          {
            x: 20,
            y: 10,
            width: 100,
            height: 20,
            baseline: 25,
            start: 0,
            end: 5,
            runs: [],
            caretStops: [
              { position: 0, x: 20 },
              { position: 5, x: 120 },
            ],
          },
          {
            x: 20,
            y: 40,
            width: 100,
            height: 20,
            baseline: 55,
            start: 5,
            end: 10,
            runs: [],
            caretStops: [
              { position: 5, x: 20 },
              { position: 10, x: 120 },
            ],
          },
        ],
      },
    ],
  };

  it('paints a shared endpoint on its visual line', () => {
    const calls: number[][] = [];
    const ctx = {
      save: () => undefined,
      restore: () => undefined,
      setTransform: () => undefined,
      fillRect: (...values: number[]) => calls.push(values),
      fillStyle: '',
    } as unknown as CanvasRenderingContext2D;

    paintSelection(
      ctx,
      frame,
      { shapeId: 'shape', storyId: 'story', anchor: 5, focus: 5, focusLine: 0 },
      1,
      1
    );

    expect(calls).toEqual([[120, 10, 1.5, 20]]);
  });

  it('stops caret blinking while blurred or hidden', async () => {
    const originalSetInterval = window.setInterval;
    const originalClearInterval = window.clearInterval;
    const visibility = Object.getOwnPropertyDescriptor(document, 'visibilityState');
    const started: number[] = [];
    const cleared: number[] = [];
    const active = new Set<number>();
    window.setInterval = ((() => {
      const id = started.length + 1;
      started.push(id);
      active.add(id);
      return id;
    }) as unknown) as typeof window.setInterval;
    window.clearInterval = (((id?: number) => {
      if (id !== undefined) {
        cleared.push(id);
        active.delete(id);
      }
    }) as unknown) as typeof window.clearInterval;
    Object.defineProperty(document, 'visibilityState', {
      configurable: true,
      value: 'visible',
    });
    const selection = {
      shapeId: 'shape',
      storyId: 'story',
      anchor: 5,
      focus: 5,
    };
    const view = render(
      <SelectionOverlay frame={frame} selection={selection} scale={1} focused />
    );

    try {
      await waitFor(() => expect(started.length).toBeGreaterThan(0));
      const focusedTimer = started[started.length - 1];
      await act(async () => {
        view.rerender(
          <SelectionOverlay frame={frame} selection={selection} scale={1} focused={false} />
        );
      });
      expect(cleared).toContain(focusedTimer);
      expect(active.size).toBe(0);

      const beforeRefocus = started.length;
      await act(async () => {
        view.rerender(
          <SelectionOverlay frame={frame} selection={selection} scale={1} focused />
        );
      });
      expect(started.length).toBeGreaterThan(beforeRefocus);
      expect(active.size).toBe(1);

      const visibleTimer = started[started.length - 1];
      Object.defineProperty(document, 'visibilityState', {
        configurable: true,
        value: 'hidden',
      });
      await act(async () => {
        fireEvent(document, new Event('visibilitychange'));
      });
      expect(cleared).toContain(visibleTimer);
      expect(active.size).toBe(0);
    } finally {
      view.unmount();
      window.setInterval = originalSetInterval;
      window.clearInterval = originalClearInterval;
      if (visibility) Object.defineProperty(document, 'visibilityState', visibility);
      else Reflect.deleteProperty(document, 'visibilityState');
    }
  });
});

describe('PptxEditor proposal review', () => {
  it('waits for the current canvas diff to paint before enabling acceptance', async () => {
    const complete: Array<() => void> = [];
    const painting = spyOn(pptx, 'paintSlide').mockImplementation((_ctx, _frame, _dpr, _scale, options) =>
      options?.textChanges ? new Promise<void>((resolve) => complete.push(resolve)) : Promise.resolve());
    const context = spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
      setTransform() {}, drawImage() {},
    } as unknown as CanvasRenderingContext2D);
    let api: PptxEditorApi | undefined;
    let view: ReturnType<typeof render> | undefined;
    try {
      view = render(<PptxEditor file={fixture} fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]}
        clientId={9212} onReady={(ready) => { api = ready; }} />);
      await act(async () => { await waitFor(() => expect(api).toBeDefined()); });
      const handle = api!.handle;
      const slideId = handle.snapshot().slides[0].id;
      await act(async () => {
        handle.propose('Review agent', null, [{ type: 'setSlideNotes', slideId, text: 'Pending notes' }]);
        api!.refreshProposals();
      });
      await waitFor(() => expect(complete.length).toBe(1));
      const accept = view.getByTestId('pptx-canvas-proposal-accept') as HTMLButtonElement;
      expect(isDisabled(accept)).toBe(true);
      expect(accept.title).toBe('The change preview is still loading.');
      await act(async () => { api!.refresh(); });
      await waitFor(() => expect(complete.length).toBe(2));
      await act(async () => { complete[0](); });
      expect(isDisabled(accept)).toBe(true);
      await act(async () => { complete[1](); });
      await waitFor(() => expect(isDisabled(accept)).toBe(false));
      fireEvent.click(accept);
      expect(handle.snapshot().slides[0].notes).toBe('Pending notes');
    } finally {
      view?.unmount();
      painting.mockRestore();
      context.mockRestore();
    }
  }, 30000);

  it('previews, accepts, undoes, rejects, and reviews stale targets through the editor UI', async () => {
    let api: PptxEditorApi | undefined;
    const view = render(<PptxEditor file={fixture} fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]}
      clientId={9210} onReady={(ready) => { api = ready; }} />);
    await act(async () => { await waitFor(() => expect(api).toBeDefined(), { timeout: 15000 }); });
    const handle = api!.handle;
    const original = handle.snapshot();
    const slide = original.slides[0];
    const shape = slide.shapes.find((shape) => shape.textStories.length > 0)!;
    const story = shape.textStories[0];
    const end = story.paragraphs[0].runs.reduce((length, run) => length + run.text.length, 0);
    await act(async () => {
      handle.propose('Review agent', 'Tighten the title', [{ type: 'replaceText', storyId: story.id, start: 0, end, text: 'A reviewed title' }]);
      api!.refreshProposals();
    });
    expect(handle.snapshot()).toEqual(original);
    expect(view.getByTestId('pptx-proposals-count').textContent).toBe('1');
    expect(view.getByTestId('pptx-canvas-proposal-diff').textContent).toContain('A reviewed title');
    fireEvent.keyDown(view.getByRole('application'), { key: 'Backspace' });
    expect(handle.snapshot()).toEqual(original);
    fireEvent.click(view.getByTestId('pptx-canvas-review-toggle'));
    expect(view.queryByTestId('pptx-canvas-proposal-diff')).toBeNull();
    fireEvent.click(view.getByTestId('pptx-canvas-review-toggle'));
    expect(view.getByTestId('pptx-canvas-proposal-diff')).toBeDefined();
    fireEvent.click(view.getByTestId('pptx-proposals-button'));
    expect(view.getByTestId('pptx-proposal').textContent).toContain('A reviewed title');
    fireEvent.click(view.getByTestId('pptx-proposal-preview'));
    await waitFor(() => expect((view.getByTestId('pptx-proposal-preview-dialog') as HTMLDialogElement).open).toBe(true));
    expect(view.getByRole('img', { name: 'Current slide' })).toBeDefined();
    expect(view.getByRole('img', { name: 'Proposed slide' })).toBeDefined();
    const dialog = view.getByTestId('pptx-proposal-preview-dialog');
    fireEvent.click(Array.from(dialog.querySelectorAll('button')).find((button) => button.textContent === 'Close')!);
    fireEvent.click(view.getByTestId('pptx-proposal-accept'));
    await waitFor(() => expect(handle.listProposals()).toHaveLength(0));
    expect(JSON.stringify(handle.snapshot())).toContain('A reviewed title');
    fireEvent.keyDown(view.getByRole('application'), { key: 'z', ...mod() });
    await waitFor(() => expect(handle.snapshot()).toEqual(original));
    await act(async () => {
      handle.propose('Review agent', null, [{ type: 'setSlideNotes', slideId: slide.id, text: 'Reject this' }]);
      api!.refreshProposals();
    });
    fireEvent.click(view.getByTestId('pptx-proposal-reject'));
    expect(handle.snapshot()).toEqual(original);
    await act(async () => {
      handle.propose('Review agent', 'Updated speaker notes', [{ type: 'setSlideNotes', slideId: slide.id, text: 'Proposed notes' }]);
      handle.setSlideNotes(slide.id, 'Human notes');
      api!.refresh();
    });
    expect(view.getByTestId('pptx-proposal-stale')).toBeDefined();
    expect(isDisabled(view.getByTestId('pptx-canvas-proposal-accept'))).toBe(true);
    expect(isDisabled(view.getByTestId('pptx-proposal-accept'))).toBe(true);
    expect(view.getByTestId('pptx-proposal-accept').title).toBe(
      "The proposal's targets changed. Review the updated preview before applying."
    );
    expect(view.getByTestId('pptx-proposal-notes-diff').textContent).toContain('Human notes');
    fireEvent.click(view.getByTestId('pptx-proposal-accept'));
    expect(handle.snapshot().slides[0].notes).toBe('Human notes');
    fireEvent.click(view.getByTestId('pptx-proposal-preview'));
    await waitFor(() => expect(view.getByTestId('pptx-proposal-force')).toBeDefined());
    handle.setSlideNotes(slide.id, 'Newer human notes');
    fireEvent.click(view.getByTestId('pptx-proposal-force'));
    expect(handle.snapshot().slides[0].notes).toBe('Newer human notes');
    expect(view.getByTestId('pptx-proposal-preview-dialog').textContent).toContain('Newer human notes');
    fireEvent.click(view.getByTestId('pptx-proposal-force'));
    await waitFor(() => expect(handle.snapshot().slides[0].notes).toBe('Proposed notes'));
    expect(handle.listProposals()).toHaveLength(0);
    await act(async () => {
      handle.propose('Review agent', 'Review both slides', [
        { type: 'setSlideNotes', slideId: original.slides[0].id, text: 'First slide' },
        { type: 'setSlideNotes', slideId: original.slides[1].id, text: 'Second slide' },
      ]);
      api!.refreshProposals();
    });
    fireEvent.click(view.getAllByTestId('pptx-proposal-preview')[1]);
    await waitFor(() => expect(view.getByTestId('pptx-proposal-preview-dialog').textContent).toContain('Second slide'));
  }, 60000);

  it('switches canvas proposals and never broadcasts a removed target selection', async () => {
    let api: PptxEditorApi | undefined;
    const cursors: Array<PptxPresenceCursor | null> = [];
    const presence = { peers: [], setCursor: (cursor: PptxPresenceCursor | null) => cursors.push(cursor), onPresence: () => () => {} };
    const view = render(<PptxEditor file={fixture} fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]}
      collaboration={{ clientId: 9211, presence }} onReady={(ready) => { api = ready; }} />);
    await act(async () => { await waitFor(() => expect(api).toBeDefined()); });
    const handle = api!.handle;
    const original = handle.snapshot();
    const slide = original.slides[0];
    const shape = slide.shapes.find((shape) => shape.textStories.length > 0)!;
    const story = shape.textStories[0];
    let first: string;
    let second: string;
    await act(async () => {
      first = handle.propose('First agent', null, [{ type: 'replaceText', storyId: story.id, start: 0, end: 0, text: 'First proposal ' }]).id;
      second = handle.propose('Second agent', null, [{ type: 'replaceText', storyId: story.id, start: 0, end: 0, text: 'Second proposal ' }]).id;
      api!.refreshProposals();
    });
    const picker = view.getByRole('combobox', { name: 'Changes on this slide' });
    expect(view.getByTestId('pptx-canvas-proposal-diff').textContent).toContain('First proposal');
    fireEvent.change(picker, { target: { value: second! } });
    expect(view.getByTestId('pptx-canvas-proposal-diff').textContent).toContain('Second proposal');
    expect(handle.snapshot()).toEqual(original);
    fireEvent.click(view.getByTestId('pptx-canvas-proposal-reject'));
    expect(handle.listProposals().map((proposal) => proposal.id)).toEqual([first!]);
    await act(async () => {
      handle.removeShape(slide.id, shape.id);
      api!.refresh();
    });
    fireEvent.click(view.getByTestId('pptx-proposals-button'));
    const card = view.getByTestId('pptx-proposal');
    const link = Array.from(card.querySelectorAll('button')).find((button) => button.textContent?.includes(shape.name))!;
    const cursorCount = cursors.length;
    fireEvent.click(link);
    await act(async () => {});
    expect(cursors.slice(cursorCount).some((cursor) => cursor?.shapeId === shape.id)).toBe(false);
    expect(isDisabled(view.getByTestId('pptx-canvas-proposal-accept'))).toBe(true);
    fireEvent.click(view.getByTestId('pptx-canvas-proposal-reject'));
    expect(view.queryByTestId('pptx-canvas-review-toolbar')).toBeNull();
  }, 30000);
});

describe('PptxEditor speaker notes', () => {
  it('saves immediately and follows undo, redo, and remote updates', async () => {
    let api: PptxEditorApi | undefined;
    const view = render(
      <PptxEditor
        file={fixture}
        fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]}
        clientId={9110}
        onReady={(ready) => {
          api = ready;
        }}
      />
    );
    await act(async () => {
      await waitFor(() => expect(api).toBeDefined(), { timeout: 15_000 });
    });
    const input = (await view.findByRole(
      'textbox',
      { name: 'Speaker notes' },
      { timeout: 15_000 }
    )) as HTMLTextAreaElement;
    const before = input.value;
    fireEvent.change(input, { target: { value: 'Remember the demo' } });
    const { openPresentation } = await import('@betteroffice/pptx');
    const saved = openPresentation(api!.save(), {
      clientId: 9111,
      fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
    });
    try {
      expect(saved.snapshot().slides[0].notes).toBe('Remember the demo');
    } finally {
      saved.dispose();
    }
    act(() => {
      api!.handle.undo();
      api!.refresh();
    });
    expect(input.value).toBe(before);
    act(() => {
      api!.handle.redo();
      api!.refresh();
    });
    expect(input.value).toBe('Remember the demo');
    const peer = openPresentation(fixture, {
      clientId: 9112,
      fonts: [{ family: 'Liberation Sans', bytes: fontBytes }],
      initialUpdate: api!.handle.encodeStateAsUpdate(),
    });
    try {
      const slide = peer.snapshot().slides[0];
      peer.setSlideNotes(slide.id, 'Remote speaker notes');
      act(() => {
        api!.handle.applyUpdate(peer.encodeStateAsUpdate());
      });
      expect(input.value).toBe('Remote speaker notes');
    } finally {
      peer.dispose();
    }
  }, 60_000);
});

describe('PptxEditor host controls', () => {
  for (const decision of [true, false, undefined] as const) {
    it(`awaits and coalesces save requests returning ${String(decision)}`, async () => {
      let api: PptxEditorApi | undefined;
      let calls = 0;
      let release!: (value: boolean | void) => void;
      const decisionPromise = new Promise<boolean | void>((resolve) => { release = resolve; });
      const saved: Uint8Array[] = [];
      const view = render(<PptxEditor file={fixture} fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]} onReady={(ready) => { api = ready; }}
        onSaveRequest={() => { calls++; return decisionPromise; }} onSave={(bytes) => saved.push(bytes)} />);
      await waitFor(() => expect(api).toBeDefined());
      if (!view.queryByTestId('pptx-save')) fireEvent.click(view.getByTestId('pptx-toolbar-more'));
      const serialize = spyOn(api!.handle, 'save');
      try {
        fireEvent.click(view.getByTestId('pptx-save'));
        fireEvent.click(view.getByTestId('pptx-save'));
        await waitFor(() => expect(calls).toBe(1));
        expect(serialize).not.toHaveBeenCalled();
        await act(async () => { release(decision); });
        expect(saved).toHaveLength(decision === true ? 1 : 0);
        expect(serialize).toHaveBeenCalledTimes(decision === true ? 1 : 0);
      } finally { serialize.mockRestore(); }
    });
  }

  it('discards an awaiting save after replacement and rejects stale flush handles', async () => {
    const opened: PptxEditorApi[] = [];
    let release!: (value: boolean) => void;
    let requests = 0;
    const pending = new Promise<boolean>((resolve) => { release = resolve; });
    const saved: Uint8Array[] = [];
    const props = { fonts: [{ family: 'Liberation Sans', bytes: fontBytes }], onReady: (ready: PptxEditorApi) => { opened.push(ready); },
      onSaveRequest: () => { requests++; return pending; }, onSave: (bytes: Uint8Array) => { saved.push(bytes); } };
    const view = render(<PptxEditor {...props} file={fixture} />);
    await waitFor(() => expect(opened).toHaveLength(1));
    if (!view.queryByTestId('pptx-save')) fireEvent.click(view.getByTestId('pptx-toolbar-more'));
    fireEvent.click(view.getByTestId('pptx-save'));
    await waitFor(() => expect(requests).toBe(1));
    view.rerender(<PptxEditor {...props} file={fixture.slice()} />);
    await waitFor(() => expect(opened).toHaveLength(2));
    await act(async () => { release(true); });
    expect(saved).toHaveLength(0);
    await expect(opened[0].flushPendingInput()).rejects.toThrow('no longer open');
    expect(() => opened[0].save()).toThrow('no longer open');
    expect(opened[0].getPositionAtPoint(1, 1)).toBeNull();
  });

  it('resolves scaled client points on the current slide without selecting or focusing', async () => {
    let api: PptxEditorApi | undefined;
    const view = render(<PptxEditor file={fixture} fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]}
      onReady={(ready) => { api = ready; }} />);
    await waitFor(() => expect(api).toBeDefined());
    const canvas = view.getByTestId('pptx-slide-canvas');
    for (const slide of [1, 2]) {
      await act(async () => { api!.goToSlide(slide); });
      const frame = api!.handle.layoutSlide(slide - 1);
      canvas.getBoundingClientRect = () => new DOMRect(40, 60, frame.width * 0.75, frame.height * 0.75);
      let point: { x: number; y: number } | undefined;
      for (let y = 0; y < frame.height && !point; y += 8) {
        for (let x = 0; x < frame.width; x += 8) {
          if (api!.handle.hitTest(x, y)?.kind === 'text') { point = { x, y }; break; }
        }
      }
      expect(point).toBeDefined();
      const snapshot = api!.handle.snapshot();
      const focused = document.activeElement;
      expect(api!.getPositionAtPoint(40 + point!.x * 0.75, 60 + point!.y * 0.75)).toEqual({
        ...api!.handle.hitTest(point!.x, point!.y)!, slide, slideId: snapshot.slides[slide - 1].id,
      });
      expect(api!.getPositionAtPoint(NaN, 70)).toBeNull();
      expect(api!.getPositionAtPoint(39, 70)).toBeNull();
      expect(api!.handle.snapshot()).toEqual(snapshot);
      expect(document.activeElement).toBe(focused);
    }
    await api!.flushPendingInput();
  });

  it('waits for accepted image input and propagates its failure', async () => {
    let api: PptxEditorApi | undefined;
    const view = render(<PptxEditor file={fixture} fonts={[{ family: 'Liberation Sans', bytes: fontBytes }]} onReady={(ready) => { api = ready; }} />);
    await waitFor(() => expect(api).toBeDefined());
    let reader!: FileReader;
    const read = spyOn(FileReader.prototype, 'readAsDataURL').mockImplementation(function (this: FileReader) { reader = this; });
    try {
      fireEvent.change(view.getByTestId('pptx-insert-image-input'), { target: { files: [new File(['png'], 'image.png', { type: 'image/png' })] } });
      let finished = false;
      const flush = api!.flushPendingInput().finally(() => { finished = true; });
      void flush.catch(() => {});
      await Promise.resolve();
      expect(finished).toBe(false);
      expect(() => api!.save()).toThrow('flushPendingInput');
      await act(async () => { reader.dispatchEvent(new Event('error')); });
      await expect(flush).rejects.toThrow();
      await api!.flushPendingInput();
      expect(api!.save().byteLength).toBeGreaterThan(0);
    } finally { read.mockRestore(); }
  });
});

describe('PptxEditor commands', () => {
  const faces = () => [{ family: 'Liberation Sans', bytes: fontBytes }];

  async function open(props: Partial<Parameters<typeof PptxEditor>[0]> = {}) {
    const opened: PptxEditorApi[] = [];
    const view = render(
      <PptxEditor file={fixture} fonts={faces()} onReady={(api) => opened.push(api)} {...props} />
    );
    await act(async () => {
      await waitFor(() => expect(opened).toHaveLength(1), { timeout: 15_000 });
    });
    const api = opened[0];
    const slide = api.handle.snapshot().slides[0];
    const shape = slide.shapes.find((candidate) => candidate.textStories.length > 0)!;
    const story = shape.textStories[0];
    return { view, api, opened, shape, story };
  }

  function pauseImages() {
    const originalImage = globalThis.Image;
    const pending: Array<{ load(): void; fail(): void }> = [];
    class PausedImage {
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      naturalWidth = 400;
      naturalHeight = 200;
      set src(value: string) {
        if (value.startsWith('data:')) {
          pending.push({ load: () => this.onload?.(), fail: () => this.onerror?.() });
        } else queueMicrotask(() => this.onload?.());
      }
    }
    globalThis.Image = PausedImage as unknown as typeof Image;
    return { pending, restore: () => (globalThis.Image = originalImage) };
  }

  function chooseImage(view: ReturnType<typeof render>) {
    const file = new File([Uint8Array.from([0x89, 0x50, 0x4e, 0x47])], 'queued.png', { type: 'image/png' });
    fireEvent.change(view.getByTestId('pptx-insert-image-input'), { target: { files: [file] } });
  }

  it('follows a presentation the host replaces from onReady', async () => {
    const opened: PptxEditorApi[] = [];
    const errors: Error[] = [];
    const fonts = faces();
    function Host() {
      const [file, setFile] = useState(fixture);
      return (
        <PptxEditor
          file={file}
          fonts={fonts}
          onError={(error) => errors.push(error)}
          onReady={(api) => {
            if (opened.push(api) === 1) setFile(new Uint8Array(fixture));
          }}
        />
      );
    }
    render(<Host />);
    await waitFor(() => expect(opened).toHaveLength(2), { timeout: 15_000 });
    const api = opened[1];
    expect(api.commands.getState('undo').enabled).toBe(false);
    const slideId = api.handle.snapshot().slides[0].id;
    await act(async () => {
      api.handle.setSlideNotes(slideId, 'After replacement');
    });
    await waitFor(() => expect(api.commands.getState('undo').enabled).toBe(true));
    expect(errors).toEqual([]);
  }, 30_000);

  it('forced acceptance checks the reviewed preview again once pending input has run', async () => {
    const images = pauseImages();
    try {
      const { view, api } = await open();
      const handle = api.handle;
      const slideId = handle.snapshot().slides[0].id;
      await act(async () => {
        handle.propose('Review agent', null, [{ type: 'setSlideNotes', slideId, text: 'Proposed notes' }]);
        handle.setSlideNotes(slideId, 'Human notes');
        api.refresh();
      });
      fireEvent.click(view.getByTestId('pptx-proposals-button'));
      chooseImage(view);
      await waitFor(() => expect(images.pending).toHaveLength(1));
      fireEvent.click(view.getByTestId('pptx-proposal-preview'));
      await waitFor(() => expect(view.getByTestId('pptx-proposal-force')).toBeDefined());
      fireEvent.click(view.getByTestId('pptx-proposal-force'));
      handle.setSlideNotes(slideId, 'Newer human notes');
      await act(async () => images.pending[0].load());
      await waitFor(() =>
        expect(view.getByTestId('pptx-proposal-preview-dialog').textContent).toContain('Newer human notes')
      );
      expect(handle.snapshot().slides[0].notes).toBe('Newer human notes');
      expect(handle.listProposals()).toHaveLength(1);
      fireEvent.click(view.getByTestId('pptx-proposal-force'));
      await waitFor(() => expect(handle.snapshot().slides[0].notes).toBe('Proposed notes'));
    } finally {
      images.restore();
    }
  }, 30_000);

  it('a deferred forced acceptance leaves a change selected meanwhile on screen', async () => {
    const images = pauseImages();
    try {
      const { view, api } = await open();
      const handle = api.handle;
      const [first, second] = handle.snapshot().slides;
      await act(async () => {
        handle.propose('Review agent', null, [
          { type: 'setSlideNotes', slideId: first.id, text: 'First proposed' },
          { type: 'setSlideNotes', slideId: second.id, text: 'Second proposed' },
        ]);
        handle.setSlideNotes(first.id, 'Human notes');
        api.refresh();
      });
      fireEvent.click(view.getByTestId('pptx-proposals-button'));
      chooseImage(view);
      await waitFor(() => expect(images.pending).toHaveLength(1));
      fireEvent.click(view.getAllByTestId('pptx-proposal-preview')[0]);
      await waitFor(() => expect(view.getByTestId('pptx-proposal-force')).toBeDefined());
      fireEvent.click(view.getByTestId('pptx-proposal-force'));
      handle.setSlideNotes(first.id, 'Newer human notes');
      const dialog = view.getByTestId('pptx-proposal-preview-dialog');
      const picker = within(dialog).getByRole('combobox', { name: 'Change' }) as HTMLSelectElement;
      fireEvent.change(picker, { target: { value: '1' } });
      await waitFor(() => expect(dialog.textContent).toContain('Second proposed'));
      const layouts = spyOn(handle, 'layoutSlide');
      await act(async () => images.pending[0].load());
      await act(async () => {
        await new Promise((done) => setTimeout(done, 50));
      });
      expect(layouts.mock.calls.map(([index]) => index)).not.toContain(0);
      expect(dialog.textContent).not.toContain('The target changed again');
      expect(picker.value).toBe('1');
      expect(handle.snapshot().slides[0].notes).toBe('Newer human notes');
      expect(handle.listProposals()).toHaveLength(1);
      layouts.mockRestore();
    } finally {
      images.restore();
    }
  }, 30_000);

  it('formats through api.commands and the platform-aware shortcuts', async () => {
    const { view, api, shape, story } = await open();
    const end = api.handle.story(story.id).length - 1;
    await act(async () => {
      api.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end });
    });
    expect(api.commands.getState('bold')).toMatchObject({ enabled: true, active: true });
    expect(view.getByTestId('pptx-bold').title).toBe('Bold (Ctrl+B)');
    const before = api.handle.snapshot();
    let result: unknown;
    await act(async () => {
      result = await api.commands.execute('bold', null);
    });
    expect(result).toEqual({ ok: true, status: 'executed' });
    expect(api.handle.story(story.id).paragraphs[0].runs[0].style.bold).toBe(false);
    expect(api.commands.getState('bold').active).toBe(false);
    fireEvent.keyDown(view.getByRole('application'), { key: 'z', ctrlKey: true });
    await waitFor(() => expect(api.handle.snapshot()).toEqual(before));
    fireEvent.keyDown(document.body, { key: 'z', ctrlKey: true, shiftKey: true });
    expect(api.handle.snapshot()).toEqual(before);
  }, 60_000);

  it('runs a command and later typing after pending image input, in order', async () => {
    const images = pauseImages();
    try {
      const { view, api, shape, story } = await open();
      await act(async () => {
        api.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 0 });
      });
      const before = api.handle.snapshot();
      const text = () =>
        api.handle
          .story(story.id)
          .paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
          .join('\n');
      const original = text();
      chooseImage(view);
      await waitFor(() => expect(images.pending).toHaveLength(1));
      const stage = view.getByRole('application');
      let undone: unknown;
      const undo = api.commands.execute('undo', null).then((value) => (undone = value));
      expect(fireEvent.keyDown(stage, { key: 'x' })).toBe(false);
      expect(fireEvent.keyDown(stage, { key: 'y' })).toBe(false);
      expect(api.handle.snapshot()).toEqual(before);
      await act(async () => {
        images.pending[0].load();
        await undo;
      });
      await api.flushPendingInput();
      expect(undone).toEqual({ ok: true, status: 'executed' });
      expect(text()).toBe(`xy${original}`);
      const shapes = api.handle.snapshot().slides[0].shapes;
      expect(shapes.some((candidate) => candidate.name === 'queued.png')).toBe(false);
      expect(shapes).toHaveLength(before.slides[0].shapes.length);
    } finally {
      images.restore();
    }
  }, 60_000);

  for (const typed of ['x', 'first character'] as const) {
    it(`carries queued typing across an undo issued between keystrokes (${typed})`, async () => {
      const images = pauseImages();
      try {
        const { view, api, shape, story } = await open();
        await act(async () => {
          api.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 0 });
        });
        const text = () =>
          api.handle
            .story(story.id)
            .paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
            .join('\n');
        const original = text();
        const key = typed === 'x' ? 'x' : original[0];
        chooseImage(view);
        await waitFor(() => expect(images.pending).toHaveLength(1));
        const stage = view.getByRole('application');
        expect(fireEvent.keyDown(stage, { key })).toBe(false);
        const undo = api.commands.execute('undo', null);
        expect(fireEvent.keyDown(stage, { key: 'y' })).toBe(false);
        await act(async () => {
          images.pending[0].load();
          await undo;
        });
        await api.flushPendingInput();
        expect(await undo).toEqual({ ok: true, status: 'executed' });
        expect(text()).toBe(`y${original}`);
      } finally {
        images.restore();
      }
    }, 60_000);
  }

  it('fails commands behind refused typing and drops input queued for a replaced deck', async () => {
    const images = pauseImages();
    const errors: Error[] = [];
    try {
      const { view, api, shape, story } = await open({ onError: (error) => errors.push(error) });
      await act(async () => {
        api.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 0 });
      });
      const refusal = new Error('typing refused');
      const insert = spyOn(api.handle, 'insertText').mockImplementation(() => {
        throw refusal;
      });
      chooseImage(view);
      await waitFor(() => expect(images.pending).toHaveLength(1));
      const stage = view.getByRole('application');
      fireEvent.keyDown(stage, { key: 'q' });
      const blocked = api.commands.execute('redo', null);
      await act(async () => {
        images.pending[0].load();
      });
      expect(await blocked).toMatchObject({ ok: false, failure: { code: 'input-failed' } });
      expect(errors).toEqual([refusal]);
      insert.mockRestore();

      chooseImage(view);
      await waitFor(() => expect(images.pending).toHaveLength(2));
      fireEvent.keyDown(stage, { key: 'z' });
      const replaced: PptxEditorApi[] = [];
      await act(async () => {
        view.rerender(
          <PptxEditor
            file={fixture.slice()}
            fonts={faces()}
            onReady={(ready) => replaced.push(ready)}
            onError={(error) => errors.push(error)}
          />
        );
      });
      await waitFor(() => expect(replaced).toHaveLength(1), { timeout: 15_000 });
      const fresh = replaced[0].handle.snapshot();
      await act(async () => {
        images.pending[1].load();
      });
      await replaced[0].flushPendingInput();
      expect(replaced[0].handle.snapshot()).toEqual(fresh);
      expect(errors).toHaveLength(1);
    } finally {
      images.restore();
    }
  }, 60_000);

  it('keeps a picker opened for one deck from inserting into its replacement', async () => {
    const originalImage = globalThis.Image;
    class FakeImage {
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      naturalWidth = 400;
      naturalHeight = 200;
      set src(_value: string) {
        queueMicrotask(() => this.onload?.());
      }
    }
    globalThis.Image = FakeImage as unknown as typeof Image;
    try {
      const { view, api } = await open();
      await act(async () => {
        expect(await api.commands.execute('insertImage', null)).toEqual({
          ok: true,
          status: 'opened',
        });
      });
      const replaced: PptxEditorApi[] = [];
      await act(async () => {
        view.rerender(
          <PptxEditor file={fixture.slice()} fonts={faces()} onReady={(ready) => replaced.push(ready)} />
        );
      });
      await waitFor(() => expect(replaced).toHaveLength(1), { timeout: 15_000 });
      const fresh = replaced[0].handle.snapshot();
      await act(async () => {
        chooseImage(view);
        await new Promise((resolve) => setTimeout(resolve, 0));
      });
      await replaced[0].flushPendingInput();
      expect(replaced[0].handle.snapshot()).toEqual(fresh);
      await act(async () => {
        chooseImage(view);
        await new Promise((resolve) => setTimeout(resolve, 0));
      });
      await waitFor(() =>
        expect(replaced[0].handle.snapshot().slides[0].shapes.length).toBe(
          fresh.slides[0].shapes.length + 1
        )
      );
    } finally {
      globalThis.Image = originalImage;
    }
  }, 60_000);

  it('evaluates consecutive commands against the state the previous one left', async () => {
    const { api, story } = await open();
    const slideId = api.handle.snapshot().slides[0].id;
    await act(async () => {
      api.handle.propose('Review agent', null, [
        { type: 'replaceText', storyId: story.id, start: 0, end: 0, text: 'Proposed ' },
      ]);
      api.refreshProposals();
    });
    expect(slideId).toBeDefined();
    const results: unknown[] = [];
    await act(async () => {
      const opening = api.commands.execute('proposalsPanel', null);
      const closing = api.commands.execute('proposalsPanel', null);
      const hiding = api.commands.execute('proposalDiff', { enabled: false });
      const showing = api.commands.execute('proposalDiff', { enabled: true });
      const zoomed = api.commands.execute('zoom', { scale: 2 });
      const again = api.commands.execute('zoom', { scale: 2 });
      results.push(...(await Promise.all([opening, closing, hiding, showing, zoomed, again])));
    });
    expect(results).toEqual([
      { ok: true, status: 'opened' },
      { ok: true, status: 'executed' },
      { ok: true, status: 'executed' },
      { ok: true, status: 'executed' },
      { ok: true, status: 'executed' },
      { ok: true, status: 'noop' },
    ]);
    expect(api.commands.getState('proposalsPanel').value).toBe(false);
    expect(api.commands.getState('proposalDiff').value).toBe(true);
    expect(api.commands.getState('zoom').value).toBe(2);
  }, 60_000);

  it('fails commands behind failed image input and read-only transitions', async () => {
    const images = pauseImages();
    const errors: Error[] = [];
    try {
      const { view, api, shape, story } = await open({ onError: (error) => errors.push(error) });
      await act(async () => {
        api.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 3 });
      });
      chooseImage(view);
      await waitFor(() => expect(images.pending).toHaveLength(1));
      const failed = api.commands.execute('italic', null);
      await act(async () => {
        images.pending[0].fail();
      });
      expect(await failed).toMatchObject({ ok: false, failure: { code: 'input-failed' } });
      expect(errors).toHaveLength(1);

      chooseImage(view);
      await waitFor(() => expect(images.pending).toHaveLength(2));
      const blocked = api.commands.execute('underline', null);
      await act(async () => {
        view.rerender(
          <PptxEditor file={fixture} fonts={faces()} onReady={() => {}} readOnly />
        );
      });
      await act(async () => {
        images.pending[1].load();
      });
      expect(await blocked).toMatchObject({ ok: false, failure: { code: 'read-only' } });
      expect(errors).toHaveLength(1);
    } finally {
      images.restore();
    }
  }, 60_000);

  it('stores caret formatting for the next typed text', async () => {
    const { view, api, shape, story } = await open();
    await act(async () => {
      api.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 0 });
    });
    expect(api.commands.getState('italic').active).toBe(false);
    const before = api.handle.snapshot();
    await act(async () => {
      expect(await api.commands.execute('italic', null)).toEqual({ ok: true, status: 'executed' });
    });
    expect(api.handle.snapshot()).toEqual(before);
    expect(api.commands.getState('italic').active).toBe(true);
    fireEvent.keyDown(view.getByRole('application'), { key: 'Q' });
    const first = api.handle.story(story.id).paragraphs[0].runs[0];
    expect(first.text.startsWith('Q')).toBe(true);
    expect(first.style.italic).toBe(true);
  }, 60_000);

  it('refuses document commands during a pointer gesture and after replacement', async () => {
    const images = pauseImages();
    try {
      const { view, api, shape, story } = await open();
      const canvas = view.getByTestId('pptx-slide-canvas');
      const frame = api.handle.layoutSlide(0);
      canvas.getBoundingClientRect = () => new DOMRect(0, 0, frame.width, frame.height);
      let point: { x: number; y: number } | undefined;
      for (let y = 0; y < frame.height && !point; y += 8) {
        for (let x = 0; x < frame.width; x += 8) {
          if (api.handle.hitTest(x, y)?.kind === 'shape') {
            point = { x, y };
            break;
          }
        }
      }
      expect(point).toBeDefined();
      canvas.setPointerCapture = () => {};
      fireEvent.pointerDown(canvas, {
        isPrimary: true,
        button: 0,
        pointerId: 7,
        clientX: point!.x,
        clientY: point!.y,
      });
      expect(await api.commands.execute('zOrder', { value: 'back' })).toMatchObject({
        ok: false,
        failure: { code: 'gesture-active' },
      });
      fireEvent.pointerUp(canvas, { pointerId: 7, clientX: point!.x, clientY: point!.y });

      await act(async () => {
        api.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 3 });
      });
      chooseImage(view);
      await waitFor(() => expect(images.pending).toHaveLength(1));
      const stale = api.commands.execute('underline', null);
      await act(async () => {
        view.rerender(
          <PptxEditor file={fixture.slice()} fonts={faces()} onReady={() => {}} />
        );
      });
      await act(async () => {
        images.pending[0].load();
      });
      expect(await stale).toMatchObject({ ok: false, failure: { code: 'document-replaced' } });
    } finally {
      images.restore();
    }
  }, 60_000);

  for (const [tool, name] of [['textBox', 'Text Box'], ['shape:rect', 'Rectangle']] as const) {
    it(`inserts a ${tool} queued behind an image decode on the slide it was drawn on`, async () => {
      const images = pauseImages();
      try {
        const { view, api } = await open();
        const before = api.handle.snapshot();
        const added = (index: number) =>
          api.handle
            .snapshot()
            .slides[index].shapes.filter((shape) => !before.slides[index].shapes.some((old) => old.id === shape.id))
            .map((shape) => shape.name);
        await act(async () => {
          await api.commands.execute('tool', { value: tool });
        });
        chooseImage(view);
        await waitFor(() => expect(images.pending).toHaveLength(1));
        const canvas = view.getByTestId('pptx-slide-canvas');
        const frame = api.handle.layoutSlide(0);
        canvas.getBoundingClientRect = () => new DOMRect(0, 0, frame.width, frame.height);
        canvas.setPointerCapture = () => {};
        canvas.hasPointerCapture = () => false;
        fireEvent.pointerDown(canvas, { isPrimary: true, button: 0, pointerId: 3, clientX: 20, clientY: 20 });
        fireEvent.pointerMove(canvas, { pointerId: 3, clientX: 140, clientY: 90 });
        fireEvent.pointerUp(canvas, { pointerId: 3, clientX: 140, clientY: 90 });
        const target = before.slides[1].shapes.find((shape) => shape.textStories.length > 0);
        expect(target).toBeDefined();
        const story = target!.textStories[0];
        await act(async () => {
          expect(
            api.selectText({ slide: 2, shapeId: target!.id, storyId: story.id, start: 0, end: 0 })
          ).toBe(true);
        });
        await waitFor(() => expect(view.container.querySelectorAll('aside canvas')).toHaveLength(before.slides.length));
        const handle = api.handle;
        const layoutSlide = handle.layoutSlide.bind(handle);
        const laidOut: number[] = [];
        handle.layoutSlide = (index) => {
          laidOut.push(index);
          return layoutSlide(index);
        };
        await act(async () => images.pending[0].load());
        await api.flushPendingInput();
        expect(added(0)).toEqual(['queued.png', name]);
        expect(added(1)).toEqual([]);
        await waitFor(() => expect(laidOut.filter((index) => index === 0).length).toBeGreaterThan(0));
        expect(laidOut.filter((index) => index === 0).length).toBeLessThanOrEqual(2);
        expect(new Set(laidOut)).toEqual(new Set([0]));
        fireEvent.keyDown(view.getByRole('application'), { key: 'Q' });
        expect(api.handle.story(story.id).paragraphs[0].runs[0].text.startsWith('Q')).toBe(true);
      } finally {
        images.restore();
      }
    }, 60_000);
  }

  it('answers shortcuts only in the owning editor, past composition, prevented keys and fields', async () => {
    const saved: Uint8Array[] = [];
    const first = await open({ onSave: (bytes) => saved.push(bytes) });
    const other: PptxEditorApi[] = [];
    const second = render(
      <PptxEditor file={fixture.slice()} fonts={faces()} onReady={(api) => other.push(api)} />
    );
    await act(async () => {
      await waitFor(() => expect(other).toHaveLength(1), { timeout: 15_000 });
    });
    await act(async () => {
      first.api.selectText({
        slide: 1,
        shapeId: first.shape.id,
        storyId: first.story.id,
        start: 0,
        end: 3,
      });
    });
    const italic = () => first.api.handle.story(first.story.id).paragraphs[0].runs[0].style.italic;
    const initial = italic();
    const untouched = other[0].handle.snapshot();
    const stage = within(first.view.container).getByRole('application');
    const press = (target: Element, init: KeyboardEventInit & { keyCode?: number } = {}) =>
      fireEvent.keyDown(target, { key: 'i', ctrlKey: true, ...init });

    press(within(second.container).getByRole('application'));
    press(document.body);
    press(stage, { isComposing: true });
    press(stage, { keyCode: 229 });
    const prevent = (event: Event) => event.preventDefault();
    stage.addEventListener('keydown', prevent, { once: true });
    press(stage);
    const notes = within(first.view.container).getByTestId('pptx-notes-textarea');
    press(notes);
    expect(italic()).toBe(initial);
    expect(other[0].handle.snapshot()).toEqual(untouched);

    expect(press(notes, { key: 's' })).toBe(false);
    await waitFor(() => expect(saved).toHaveLength(1));
    expect(press(stage)).toBe(false);
    await waitFor(() => expect(italic()).toBe(true));
  }, 60_000);

  it('owns shortcuts inside its portalled popups and keeps keyboard focus on the toolbar', async () => {
    const originalRect = HTMLElement.prototype.getBoundingClientRect;
    HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
      const rect = originalRect.call(this);
      return this.getAttribute('role') === 'toolbar' ? { ...rect.toJSON(), width: 4000 } : rect;
    };
    try {
      const { view, api, shape, story } = await open();
      await act(async () => {
        api.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 3 });
      });
      const run = () => api.handle.story(story.id).paragraphs[0].runs[0].style;
      const initialUnderline = run().underline;
      const family = view.getByTestId('pptx-font-family');
      act(() => family.focus());
      fireEvent.keyDown(family, { key: 'ArrowDown' });
      const menu = within(document.body).getByRole('menu', { name: 'Font family' });
      expect(view.container.contains(menu)).toBe(false);
      const item = document.activeElement as HTMLElement;
      expect(menu.contains(item)).toBe(true);
      fireEvent.keyDown(item, { key: 'u', ctrlKey: true });
      await waitFor(() => expect(run().underline).not.toBe(initialUnderline));
      fireEvent.keyDown(item, { key: 'Escape' });
      expect(document.activeElement).toBe(family);

      const italic = view.getByTestId('pptx-italic');
      act(() => italic.focus());
      fireEvent.keyDown(italic, { key: 'Enter' });
      fireEvent.click(italic);
      await waitFor(() => expect(run().italic).toBe(true));
      fireEvent.keyDown(italic, { key: 'b', ctrlKey: true });
      await waitFor(() => expect(run().bold).toBe(false));
      expect(document.activeElement).toBe(italic);
    } finally {
      HTMLElement.prototype.getBoundingClientRect = originalRect;
    }
  }, 60_000);

  it('lets a save request flush input and reports its outcome', async () => {
    const saved: Uint8Array[] = [];
    let decision = true;
    let api: PptxEditorApi | undefined;
    const context = await open({
      onSave: (bytes) => saved.push(bytes),
      onSaveRequest: async () => {
        await api!.flushPendingInput();
        return decision;
      },
    });
    api = context.api;
    expect(await api.commands.execute('save', null)).toEqual({ ok: true, status: 'executed' });
    decision = false;
    expect(await api.commands.execute('save', null)).toEqual({ ok: true, status: 'requested' });
    expect(saved).toHaveLength(1);
  }, 60_000);

  it('replaces, hides and keeps the toolbar region by precedence', async () => {
    const { view } = await open({ toolbar: null });
    expect(view.queryByTestId('pptx-editor-toolbar')).toBeNull();
    expect(view.queryByTestId('pptx-present')).toBeNull();
    const custom = (
      <EditorToolbar mode="commands">
        <EditorToolbar.Toolbar>
          <ToolbarCommandButton id="bold" />
          <ToolbarCommandButton id="slideshow" />
        </EditorToolbar.Toolbar>
      </EditorToolbar>
    );
    await act(async () => {
      view.rerender(
        <PptxEditor file={fixture} fonts={faces()} toolbar={custom} readOnly />
      );
    });
    expect(view.getByTestId('pptx-bold').title).toContain('The presentation is read-only.');
    expect(view.queryByTestId('pptx-save')).toBeNull();
    await act(async () => {
      view.rerender(
        <PptxEditor file={fixture} fonts={faces()} toolbar={custom} showToolbar={false} />
      );
    });
    expect(view.queryByTestId('pptx-bold')).toBeNull();
    expect(view.getByTestId('pptx-insert-image-input')).toBeDefined();
    await act(async () => {
      view.rerender(<PptxEditor file={fixture} fonts={faces()} readOnly />);
    });
    expect(view.queryByTestId('pptx-editor-toolbar')).toBeNull();
    expect(view.getByTestId('pptx-present')).toBeDefined();
  }, 60_000);

  it('drives an external toolbar and owns shortcuts pressed inside it', async () => {
    const { view, api, shape, story } = await open({ toolbar: null });
    await act(async () => {
      api.selectText({ slide: 1, shapeId: shape.id, storyId: story.id, start: 0, end: 3 });
    });
    const outside = render(
      <PptxCommandProvider commands={api.commands}>
        <EditorToolbar mode="commands">
          <EditorToolbar.Toolbar>
            <ToolbarCommandButton id="italic" />
            <ToolbarCommandButton id="undo" />
          </EditorToolbar.Toolbar>
        </EditorToolbar>
      </PptxCommandProvider>,
      { container: document.body.appendChild(document.createElement('div')) }
    );
    const italic = outside.getAllByTestId('pptx-italic').find((button) => !view.container.contains(button))!;
    const before = api.handle.snapshot();
    await act(async () => {
      fireEvent.click(italic);
    });
    await waitFor(() =>
      expect(api.handle.story(story.id).paragraphs[0].runs[0].style.italic).toBe(true)
    );
    fireEvent.keyDown(italic, { key: 'z', ctrlKey: true });
    await waitFor(() => expect(api.handle.snapshot()).toEqual(before));
    outside.unmount();
  }, 60_000);
});

describe('PptxEditor edit batches', () => {
  const fonts = () => [{ family: 'Liberation Sans', bytes: fontBytes }];

  /** Image decoding that finishes only when the test says so, keeping the insertion pending. */
  function pauseImages(): { started: () => boolean; finish: () => void; restore: () => void } {
    const originalImage = globalThis.Image;
    let pending: (() => void) | undefined;
    class PausedImage {
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;
      naturalWidth = 400;
      naturalHeight = 200;
      set src(value: string) {
        if (value.startsWith('data:')) pending = () => this.onload?.();
        else queueMicrotask(() => this.onload?.());
      }
    }
    globalThis.Image = PausedImage as unknown as typeof Image;
    return {
      started: () => pending !== undefined,
      finish: () => pending?.(),
      restore: () => {
        globalThis.Image = originalImage;
      },
    };
  }

  const insertImage = (view: ReturnType<typeof render>, name: string) =>
    fireEvent.change(view.getByTestId('pptx-insert-image-input'), {
      target: { files: [new File([Uint8Array.from([0x89, 0x50, 0x4e, 0x47])], name, { type: 'image/png' })] },
    });

  const notesRequest = (version: string, slideId: string, text: string): PptxEditRequest => ({
    expectVersion: version,
    steps: [{ op: 'setSlideNotes', target: { slideId }, text }],
  });

  it('flushes pending input first and refuses a stale batch without undoing that input', async () => {
    const images = pauseImages();
    try {
      let api: PptxEditorApi | undefined;
      const view = render(
        <PptxEditor file={fixture} fonts={fonts()} clientId={9130} onReady={(ready) => { api = ready; }} />
      );
      await act(async () => {
        await waitFor(() => expect(api).toBeDefined(), { timeout: 15_000 });
      });
      const read = await api!.readContent();
      if (!read.ok) throw new Error(read.failure.message);
      insertImage(view, 'flushed.png');
      await waitFor(() => expect(images.started()).toBe(true));
      let settled = false;
      const applying = api!
        .applyEdits(notesRequest(read.version, read.slides[0].id, 'Too late'))
        .finally(() => { settled = true; });
      await Promise.resolve();
      expect(settled).toBe(false);
      let result: PptxEditResult | undefined;
      await act(async () => {
        images.finish();
        result = await applying;
      });
      expect(result).toMatchObject({ ok: false, failure: { code: 'stale-version' } });
      expect(result!.version).toBe(api!.handle.version());
      const slide = api!.handle.snapshot().slides[0];
      expect(slide.shapes.some((shape) => shape.name === 'flushed.png')).toBe(true);
      expect(slide.notes ?? '').not.toBe('Too late');
    } finally {
      cleanup();
      images.restore();
    }
  }, 30_000);

  it('applies a batch as one undo step and publishes it once', async () => {
    let api: PptxEditorApi | undefined;
    const changes: DeckSnapshot[] = [];
    const view = render(
      <PptxEditor file={fixture} fonts={fonts()} clientId={9131} onReady={(ready) => { api = ready; }}
        onChange={(snapshot) => changes.push(snapshot)} />
    );
    await act(async () => {
      await waitFor(() => expect(api).toBeDefined(), { timeout: 15_000 });
    });
    const read = await api!.readContent();
    if (!read.ok) throw new Error(read.failure.message);
    const slideId = read.slides[0].id;
    let result: PptxEditResult | undefined;
    await act(async () => {
      result = await api!.applyEdits(notesRequest(read.version, slideId, 'Batch notes'));
    });
    expect(result).toMatchObject({ ok: true, applied: true, changedSlides: [slideId] });
    expect(changes).toHaveLength(1);
    expect(changes[0].slides[0].notes).toBe('Batch notes');
    const notes = view.getByRole('textbox', { name: 'Speaker notes' }) as HTMLTextAreaElement;
    expect(notes.value).toBe('Batch notes');
    expect(api!.handle.canUndo()).toBe(true);
    expect(await api!.version()).toBe(api!.handle.version());
    await act(async () => {
      result = await api!.applyEdits(notesRequest(api!.handle.version(), slideId, 'Batch notes'));
    });
    expect(result).toMatchObject({ ok: true, applied: false });
    expect(changes).toHaveLength(1);
  }, 30_000);

  it('refuses writes while read-only but still reads', async () => {
    let api: PptxEditorApi | undefined;
    render(<PptxEditor file={fixture} fonts={fonts()} clientId={9132} onReady={(ready) => { api = ready; }} readOnly />);
    await act(async () => {
      await waitFor(() => expect(api).toBeDefined(), { timeout: 15_000 });
    });
    const read = await api!.readContent();
    if (!read.ok) throw new Error(read.failure.message);
    const found = await api!.findText({ text: 'e', limit: 1 });
    expect(found).toMatchObject({ ok: true, version: read.version });
    const request = notesRequest(read.version, read.slides[0].id, 'Blocked');
    const refusal = { ok: false, version: read.version, failure: { code: 'read-only' } };
    expect(await api!.validateEdits(request)).toMatchObject(refusal);
    expect(await api!.applyEdits(request)).toMatchObject(refusal);
    expect(api!.handle.snapshot().slides[0].notes ?? '').not.toBe('Blocked');
  }, 30_000);

  it('refuses a batch when the editor turns read-only while input flushes', async () => {
    const images = pauseImages();
    try {
      let api: PptxEditorApi | undefined;
      const props = { file: fixture, fonts: fonts(), clientId: 9134, onReady: (ready: PptxEditorApi) => { api = ready; } };
      const view = render(<PptxEditor {...props} />);
      await act(async () => {
        await waitFor(() => expect(api).toBeDefined(), { timeout: 15_000 });
      });
      const read = await api!.readContent();
      if (!read.ok) throw new Error(read.failure.message);
      insertImage(view, 'pending.png');
      await waitFor(() => expect(images.started()).toBe(true));
      const applying = api!.applyEdits(notesRequest(read.version, read.slides[0].id, 'Blocked'));
      await act(async () => { view.rerender(<PptxEditor {...props} readOnly />); });
      await act(async () => { images.finish(); });
      expect(await applying).toMatchObject({
        ok: false,
        version: api!.handle.version(),
        failure: { code: 'read-only' },
      });
      expect(api!.handle.snapshot().slides[0].notes ?? '').not.toBe('Blocked');
      expect(api!.handle.canUndo()).toBe(false);
    } finally {
      cleanup();
      images.restore();
    }
  }, 30_000);

  it('rejects when the presentation is replaced while input flushes', async () => {
    const images = pauseImages();
    try {
      const opened: PptxEditorApi[] = [];
      const props = { fonts: fonts(), clientId: 9133, onReady: (ready: PptxEditorApi) => { opened.push(ready); } };
      const view = render(<PptxEditor {...props} file={fixture} />);
      await act(async () => {
        await waitFor(() => expect(opened).toHaveLength(1), { timeout: 15_000 });
      });
      const read = await opened[0].readContent();
      if (!read.ok) throw new Error(read.failure.message);
      insertImage(view, 'pending.png');
      await waitFor(() => expect(images.started()).toBe(true));
      const applying = opened[0].applyEdits(notesRequest(read.version, read.slides[0].id, 'Orphaned'));
      void applying.catch(() => {});
      view.rerender(<PptxEditor {...props} file={new Uint8Array(fixture)} />);
      await act(async () => {
        await waitFor(() => expect(opened).toHaveLength(2), { timeout: 15_000 });
      });
      await act(async () => { images.finish(); });
      await expect(applying).rejects.toThrow('while flushing');
      expect(opened[1].handle.snapshot().slides[0].notes ?? '').not.toBe('Orphaned');
    } finally {
      cleanup();
      images.restore();
    }
  }, 30_000);
});
