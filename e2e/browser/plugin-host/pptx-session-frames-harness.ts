import {
  decodePresentationImage,
  initWasm,
  openPresentation,
  openPresentationSession,
  paintSlide,
  sizeCanvasForSlide,
  type CanvasImageResolver,
  type PresentationSession,
  type SlideDisplayList,
} from '../../../packages/pptx/src/index';

export interface SessionFrameResult {
  name: string;
  fixture: string;
  slideIndex: number;
  dpr: number;
  zoom: number;
  differingPixels: number;
  maxChannelDelta: number;
  versionMatches: boolean;
  mainImagesDrawn: number;
  sessionImagesDrawn: number;
}

export interface SessionFramesProbe {
  done: Promise<SessionFrameResult[]>;
  workers: { fixture: string; url: string; type: WorkerOptions['type'] }[];
}

const fixtures = [
  {
    name: 'demo',
    url: new URL('../../../apps/demo/public/betteroffice-demo.pptx', import.meta.url),
  },
  {
    name: 'tiff',
    url: new URL('../../../packages/pptx/src/render/fixtures/tiff-image.pptx', import.meta.url),
  },
];
const scales = [
  { dpr: 1, zoom: 1 },
  { dpr: 2, zoom: 1 },
  { dpr: 1, zoom: 1.5 },
];
const workers: SessionFramesProbe['workers'] = [];
const imageError = 'could not decode slide image';

function imageCache() {
  const pending = new Map<string, Promise<CanvasImageSource | null>>();
  const sources = new Set<CanvasImageSource>();
  const resolver = (mediaBytes: (assetId: string) => Uint8Array | undefined): CanvasImageResolver =>
    (assetId) => {
      const cached = pending.get(assetId);
      if (cached) return cached;
      const bytes = mediaBytes(assetId);
      const decoded = bytes ? decodePresentationImage(bytes, imageError) : Promise.resolve(null);
      const image = decoded.then((source) => {
        if (source) sources.add(source);
        return source;
      });
      pending.set(assetId, image);
      return image;
    };
  return { pending, sources, resolver };
}

async function paint(
  canvas: HTMLCanvasElement,
  displayList: SlideDisplayList,
  dpr: number,
  zoom: number,
  images: ReturnType<typeof imageCache>,
  mediaBytes: (assetId: string) => Uint8Array | undefined
) {
  sizeCanvasForSlide(canvas, displayList, dpr, zoom);
  const context = canvas.getContext('2d');
  if (!context) throw new Error('Canvas2D is unavailable');
  let imagesDrawn = 0;
  const methods = new Map<PropertyKey, (...args: unknown[]) => unknown>();
  const counted = new Proxy(context, {
    get(target, key) {
      const value: unknown = Reflect.get(target, key, target);
      if (typeof value !== 'function') return value;
      let method = methods.get(key);
      if (!method) {
        method = (...args: unknown[]) => {
          const result = Reflect.apply(value, target, args);
          if (key === 'drawImage' && images.sources.has(args[0] as CanvasImageSource)) {
            imagesDrawn += 1;
          }
          return result;
        };
        methods.set(key, method);
      }
      return method;
    },
    set: (target, key, value) => Reflect.set(target, key, value, target),
  });
  await paintSlide(counted, displayList, dpr, zoom, { resolveImage: images.resolver(mediaBytes) });
  await Promise.all(images.pending.values());
  return { context, imagesDrawn };
}

async function run(): Promise<SessionFrameResult[]> {
  const NativeWorker = globalThis.Worker;
  let activeFixture = '';
  globalThis.Worker = class extends NativeWorker {
    constructor(url: string | URL, options?: WorkerOptions) {
      super(url, options);
      workers.push({ fixture: activeFixture, url: String(url), type: options?.type });
    }
  };
  try {
    await initWasm();
    await document.fonts.ready;
    const fontResponse = await fetch(
      new URL('../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf', import.meta.url)
    );
    if (!fontResponse.ok) throw new Error(`font fetch failed: ${fontResponse.status}`);
    const fontBytes = new Uint8Array(await fontResponse.arrayBuffer());
    const fonts = [{ family: 'Liberation Sans', bytes: fontBytes }];
    const mainCanvas = document.querySelector<HTMLCanvasElement>('#main')!;
    const sessionCanvas = document.querySelector<HTMLCanvasElement>('#session')!;
    const results: SessionFrameResult[] = [];
    for (const fixture of fixtures) {
      activeFixture = fixture.name;
      const response = await fetch(fixture.url);
      if (!response.ok) throw new Error(`${fixture.name} fetch failed: ${response.status}`);
      const bytes = new Uint8Array(await response.arrayBuffer());
      const handle = openPresentation(bytes, { fonts });
      let session: PresentationSession | undefined;
      try {
        session = await openPresentationSession(bytes, { fonts });
        const slideCount = handle.snapshot().slides.length;
        if (session.state.slides.length !== slideCount) throw new Error('Slide counts differ');
        const mainImages = imageCache();
        const sessionImages = imageCache();
        for (let slideIndex = 0; slideIndex < slideCount; slideIndex += 1) {
          for (const { dpr, zoom } of scales) {
            const displayList = handle.layoutSlide(slideIndex);
            const frame = await session.call.frame(slideIndex);
            if (frame.slideIndex !== slideIndex) throw new Error('Frame slide index differs');
            const mainPaint = await paint(
              mainCanvas, displayList, dpr, zoom, mainImages, (assetId) => handle.mediaBytes(assetId)
            );
            const sessionPaint = await paint(
              sessionCanvas, frame.displayList, dpr, zoom, sessionImages,
              (assetId) => frame.media.get(assetId)
            );
            if (mainCanvas.width !== sessionCanvas.width || mainCanvas.height !== sessionCanvas.height) {
              throw new Error('Frame canvas sizes differ');
            }
            const a = mainPaint.context.getImageData(0, 0, mainCanvas.width, mainCanvas.height).data;
            const b = sessionPaint.context.getImageData(
              0, 0, sessionCanvas.width, sessionCanvas.height
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
              name: `${fixture.name} slide=${slideIndex} dpr=${dpr} zoom=${zoom}`,
              fixture: fixture.name,
              slideIndex,
              dpr,
              zoom,
              differingPixels,
              maxChannelDelta,
              versionMatches: frame.version === (await session.call.version()),
              mainImagesDrawn: mainPaint.imagesDrawn,
              sessionImagesDrawn: sessionPaint.imagesDrawn,
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

(window as unknown as { __pptxSessionFrames: SessionFramesProbe }).__pptxSessionFrames = {
  workers,
  done: run(),
};
