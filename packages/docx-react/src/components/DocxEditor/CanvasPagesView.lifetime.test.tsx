import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, expect, mock, spyOn, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import type { DisplayList, GlyphOutlineProvider } from '@betteroffice/docx/layout/render';
import { CanvasPagesView } from './CanvasPagesView';

const { act, cleanup, render } = await import('@testing-library/react');

afterEach(() => {
  cleanup();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function canvasContext() {
  return {
    resetTransform() {},
    scale() {},
    clearRect() {},
    save() {},
    restore() {},
    drawImage: mock((..._args: unknown[]) => {}),
    fillText: mock((..._args: unknown[]) => {}),
  };
}

test.each(['unmount', 'provider replacement'] as const)(
  'image-delayed canvas replay never reads retired glyph outlines after %s',
  async (retire) => {
    const contexts = new Map<HTMLCanvasElement, ReturnType<typeof canvasContext>>();
    const getContext = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
      function (this: HTMLCanvasElement) {
        let context = contexts.get(this);
        if (!context) {
          context = canvasContext();
          contexts.set(this, context);
        }
        return context as unknown as CanvasRenderingContext2D;
      } as unknown as typeof HTMLCanvasElement.prototype.getContext
    );
    const image = document.createElement('canvas');
    let finishDecode!: (image: CanvasImageSource) => void;
    const pendingImage = new Promise<CanvasImageSource>((resolve) => {
      finishDecode = resolve;
    });
    const resolveImage = mock((_relId: string) => pendingImage);
    const originalProvider = mock((_fontId: number, _glyphId: number) =>
      JSON.stringify({ upem: 1000, cmds: [] })
    );
    const replacementProvider = mock((_fontId: number, _glyphId: number) =>
      JSON.stringify({ upem: 1000, cmds: [] })
    );
    const displayList: DisplayList = {
      pages: [
        {
          pageIndex: 0,
          width: 100,
          height: 100,
          primitives: [
            { kind: 'image', relId: 'held-image', x: 0, y: 0, w: 10, h: 10 },
            {
              kind: 'glyphRun',
              fontId: 7,
              size: 12,
              color: '#000000',
              text: 'A',
              glyphs: [{ id: 42, x: 10, y: 20, cluster: 0, advance: 8 }],
            },
          ],
        },
      ],
    };
    const view = (provider: GlyphOutlineProvider) => (
      <CanvasPagesView
        displayList={displayList}
        resolveImage={resolveImage}
        glyphOutlineProvider={provider}
      />
    );

    try {
      const { rerender, unmount } = render(view(originalProvider));
      await act(async () => {});
      expect(resolveImage).toHaveBeenCalledTimes(1);
      expect(resolveImage).toHaveBeenCalledWith('held-image');
      expect(originalProvider).not.toHaveBeenCalled();
      const pendingBuffers = [...contexts].filter(([canvas]) => !canvas.isConnected);
      expect(pendingBuffers).toHaveLength(1);
      for (const [, context] of pendingBuffers) {
        expect(context.drawImage).not.toHaveBeenCalled();
        expect(context.fillText).not.toHaveBeenCalled();
      }

      if (retire === 'unmount') unmount();
      else rerender(view(replacementProvider));
      await act(async () => {});
      await act(async () => {
        finishDecode(image);
        await pendingImage;
      });

      expect(originalProvider).not.toHaveBeenCalled();
      for (const [, context] of pendingBuffers) {
        expect(context.drawImage).toHaveBeenCalledWith(image, 0, 0, 10, 10);
        expect(context.fillText).toHaveBeenCalledWith('A', 10, 20);
      }
      if (retire === 'provider replacement') {
        expect(replacementProvider).toHaveBeenCalledWith(7, 42);
      }
    } finally {
      cleanup();
      finishDecode(image);
      await act(async () => {});
      getContext.mockRestore();
    }
  }
);

test("an engine's glyph outlines paint its first replay, rasterized once", async () => {
  const contexts = new Map<HTMLCanvasElement, ReturnType<typeof canvasContext>>();
  const getContext = spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(
    function (this: HTMLCanvasElement) {
      let context = contexts.get(this);
      if (!context) {
        context = canvasContext();
        contexts.set(this, context);
      }
      return context as unknown as CanvasRenderingContext2D;
    } as unknown as typeof HTMLCanvasElement.prototype.getContext
  );
  const provider = mock((_fontId: number, _glyphId: number) =>
    JSON.stringify({ upem: 1000, cmds: [] })
  );
  const displayList: DisplayList = {
    pages: [
      {
        pageIndex: 0,
        width: 100,
        height: 100,
        primitives: [
          {
            kind: 'glyphRun',
            fontId: 7,
            size: 12,
            color: '#000000',
            text: 'A',
            glyphs: [{ id: 42, x: 10, y: 20, cluster: 0, advance: 8 }],
          },
        ],
      },
    ],
  };
  try {
    render(<CanvasPagesView displayList={displayList} glyphOutlineProvider={provider} />);
    await act(async () => {});
    const buffers = [...contexts].filter(([canvas]) => !canvas.isConnected);
    expect(buffers).toHaveLength(1);
    expect(provider).toHaveBeenCalledWith(7, 42);
    expect(buffers[0]![1].fillText).not.toHaveBeenCalled();
  } finally {
    cleanup();
    getContext.mockRestore();
  }
});
