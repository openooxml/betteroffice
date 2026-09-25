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
import * as pptx from '@betteroffice/pptx';
import type {
  DeckSnapshot,
  PptxEditRequest,
  PptxEditResult,
  PptxFontFace,
  PptxPresenceCursor,
  SlideDisplayList,
} from '@betteroffice/pptx';
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
    'lands on the slide showing when the read finishes, not the one showing when it started',
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
        await waitFor(() => expect(hasLogo(1)).toBe(true));
        expect(hasLogo(0)).toBe(false);
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
      expect(accept.disabled).toBe(true);
      await act(async () => { api!.refresh(); });
      await waitFor(() => expect(complete.length).toBe(2));
      await act(async () => { complete[0](); });
      expect(accept.disabled).toBe(true);
      await act(async () => { complete[1](); });
      await waitFor(() => expect(accept.disabled).toBe(false));
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
    fireEvent.keyDown(view.getByRole('application'), { key: 'z', ctrlKey: true, metaKey: true });
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
    expect((view.getByTestId('pptx-canvas-proposal-accept') as HTMLButtonElement).disabled).toBe(true);
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
    expect((view.getByTestId('pptx-canvas-proposal-accept') as HTMLButtonElement).disabled).toBe(true);
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
