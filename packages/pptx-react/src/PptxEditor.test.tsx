/**
 * `fonts` identity drives the effect that disposes and reopens the presentation,
 * so a caller building the array inline must not lose the open deck — its edits,
 * history and collaboration replica — on every unrelated re-render.
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, describe, expect, it, spyOn } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { initWasm } from '@betteroffice/pptx';
import type { PptxFontFace, SlideDisplayList } from '@betteroffice/pptx';
import type { PptxEditorApi } from './PptxEditor';
import { paintSelection, PptxEditor, SelectionOverlay } from './PptxEditor';

const root = resolve(import.meta.dir, '../../..');

// the registrator writes one process-wide global set, so only the file that
// installed it may tear it down.
const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
const { act, cleanup, fireEvent, render, waitFor } = await import('@testing-library/react');

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
          if (!view.queryByTestId('pptx-export-png')) {
            fireEvent.click(view.getByTestId('pptx-toolbar-more'));
          }
          fireEvent.click(view.getByTestId('pptx-export-png'));
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
