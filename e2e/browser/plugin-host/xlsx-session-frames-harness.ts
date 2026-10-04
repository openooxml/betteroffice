import {
  initWasm,
  openWorkbook,
  openWorkbookSession,
  paintDisplayList,
  type Viewport,
  type WorkbookSession,
} from '../../../packages/xlsx/src/index';

export interface SessionFrameResult {
  name: string;
  differingPixels: number;
  maxChannelDelta: number;
  versionMatches: boolean;
}

export interface SessionFramesProbe {
  done: Promise<SessionFrameResult[]>;
  workers: { url: string; type: WorkerOptions['type'] }[];
}

const fixtures = [
  {
    name: 'sample',
    url: new URL('../../../packages/xlsx/test-fixtures/sample.xlsx', import.meta.url),
  },
  {
    name: 'charts',
    url: new URL('../../../packages/xlsx/test-fixtures/charts.xlsx', import.meta.url),
  },
];
const viewports = [
  { name: 'origin', viewport: { x: 0, y: 0, width: 1200, height: 700 } },
  { name: 'scrolled', viewport: { x: 300, y: 400, width: 900, height: 500 } },
];
const scales = [
  { dpr: 1, zoom: 1 },
  { dpr: 2, zoom: 1 },
  { dpr: 1, zoom: 1.5 },
];
const workers: SessionFramesProbe['workers'] = [];

function sizeCanvas(canvas: HTMLCanvasElement, viewport: Viewport, dpr: number, zoom: number) {
  const w = viewport.width * zoom;
  const h = viewport.height * zoom;
  canvas.width = Math.round(w * dpr);
  canvas.height = Math.round(h * dpr);
  canvas.style.width = `${w}px`;
  canvas.style.height = `${h}px`;
  const context = canvas.getContext('2d');
  if (!context) throw new Error('Canvas2D is unavailable');
  return context;
}

async function run(): Promise<SessionFrameResult[]> {
  const NativeWorker = globalThis.Worker;
  globalThis.Worker = class extends NativeWorker {
    constructor(url: string | URL, options?: WorkerOptions) {
      super(url, options);
      workers.push({ url: String(url), type: options?.type });
    }
  };
  try {
    await initWasm();
    await document.fonts.ready;
    const mainCanvas = document.querySelector<HTMLCanvasElement>('#main')!;
    const sessionCanvas = document.querySelector<HTMLCanvasElement>('#session')!;
    const results: SessionFrameResult[] = [];
    for (const fixture of fixtures) {
      const response = await fetch(fixture.url);
      if (!response.ok) throw new Error(`${fixture.name} fetch failed: ${response.status}`);
      const bytes = new Uint8Array(await response.arrayBuffer());
      const handle = openWorkbook(bytes);
      let session: WorkbookSession | undefined;
      try {
        session = await openWorkbookSession(bytes);
        for (const { name, viewport } of viewports) {
          for (const { dpr, zoom } of scales) {
            const mainContext = sizeCanvas(mainCanvas, viewport, dpr, zoom);
            const sessionContext = sizeCanvas(sessionCanvas, viewport, dpr, zoom);
            const displayList = handle.displayList(viewport);
            const frame = await session.call.frame(viewport);
            paintDisplayList(mainContext, displayList, dpr * zoom);
            paintDisplayList(sessionContext, frame.displayList, dpr * zoom);
            const a = mainContext.getImageData(0, 0, mainCanvas.width, mainCanvas.height).data;
            const b = sessionContext.getImageData(
              0,
              0,
              sessionCanvas.width,
              sessionCanvas.height
            ).data;
            let differingPixels = 0;
            let maxChannelDelta = 0;
            for (let pixel = 0; pixel < a.length; pixel += 4) {
              let differs = false;
              for (let channel = 0; channel < 4; channel += 1) {
                const delta = Math.abs(a[pixel + channel]! - b[pixel + channel]!);
                if (delta !== 0) differs = true;
                maxChannelDelta = Math.max(maxChannelDelta, delta);
              }
              if (differs) differingPixels += 1;
            }
            results.push({
              name: `${fixture.name} ${name} dpr=${dpr} zoom=${zoom}`,
              differingPixels,
              maxChannelDelta,
              versionMatches: frame.version === (await session.call.version()),
            });
          }
        }
      } finally {
        try {
          await session?.dispose();
        } finally {
          handle.dispose();
        }
      }
    }
    return results;
  } finally {
    globalThis.Worker = NativeWorker;
  }
}

(window as unknown as { __xlsxSessionFrames: SessionFramesProbe }).__xlsxSessionFrames = {
  workers,
  done: run(),
};
