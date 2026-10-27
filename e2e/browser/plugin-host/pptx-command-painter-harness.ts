import { paintSlide, sizeCanvasForSlide, type SlideDisplayList, type SlidePrimitive } from '@betteroffice/pptx';
import { createCommandFramePainter } from '../../../packages/pptx-react/src/worker/commandFramePainter';
import type { PeerlessViewSnapshot } from '../../../packages/pptx-react/src/worker/peerlessViewStore';

export interface CommandPainterProbe {
  pixels(primitive: SlidePrimitive, region: { x: number; y: number; width: number; height: number }, scale?: number):
    Promise<{ actual: number[]; reference: number[] }>;
}

declare global {
  interface Window { __pptxCommandPainter: CommandPainterProbe }
}

const font = new FontFace('CommandPainterMono',
  `url("${new URL('../../../packages/fonts/assets/LiberationMono-Regular.ttf', import.meta.url).href}")`);
document.fonts.add(font);
await font.load();
await document.fonts.ready;

window.__pptxCommandPainter = {
  async pixels(primitive, region, scale = 1) {
    const list: SlideDisplayList = { contractVersion: 1, width: 960, height: 540,
      background: { kind: 'solid', color: '#ffffff' }, primitives: [primitive] };
    const canvas = document.createElement('canvas');
    const reference = document.createElement('canvas');
    const painter = createCommandFramePainter(canvas, { scale });
    try {
      const frame = { displayList: list, frame: {
        slideId: 'slide', slideIndex: 0, slideRevision: 0, coversThroughSeq: 0,
        documentEpoch: 1, incarnation: 'worker', epoch: 1, version: 'v0', media: [],
        displayList: new TextEncoder().encode(JSON.stringify(list)).buffer,
      } };
      const view: PeerlessViewSnapshot = { confirmed: frame, selection: null, focused: false, caret: null,
        selectionRects: [], composition: null, echo: null, activeLine: null, activeContextBytes: 0,
        pending: [], waitingInputCount: 0, appliedThroughSeq: 0, presentedThroughSeq: 0,
        editState: { readOnly: false, canUndo: false, canRedo: false,
          toolbar: { target: null, storedMarks: {}, formatting: null, alignment: null,
            hyperlink: null, table: null, canEditText: true } },
        thumbnails: new Map(), dirtyThumbnailIds: new Set<string>() };
      painter.paintFrame(frame, view);
      if (!await painter.whenPainted()) throw new Error('Painter did not finish');
      sizeCanvasForSlide(reference, list, devicePixelRatio, scale);
      const ctx = reference.getContext('2d');
      if (!ctx) throw new Error('Reference canvas is unavailable');
      await paintSlide(ctx, list, devicePixelRatio, scale);
      const read = (surface: HTMLCanvasElement) => {
        const ctx = surface.getContext('2d');
        if (!ctx) throw new Error('Canvas is unavailable');
        const factor = devicePixelRatio * scale;
        return Array.from(ctx.getImageData(Math.round(region.x * factor), Math.round(region.y * factor),
          Math.round(region.width * factor), Math.round(region.height * factor)).data);
      };
      return { actual: read(canvas), reference: read(reference) };
    } finally { painter.dispose(); }
  },
};
