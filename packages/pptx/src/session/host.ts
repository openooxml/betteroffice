import {
  createSessionHost,
  transferable,
  type MethodHandlers,
  type SessionHost,
  type SessionTransport,
} from '../../../../shared/office-session';
import { MAX_TIFF_BYTES, isTiff } from '../../../../shared/media';
import type { DeckSnapshot, PptxFontFace, SlideDisplayList } from '../types';
import {
  decodeTiffImage, initWasm, openPresentation, presentationDisplayListJson, type PresentationHandle,
} from '../wasm/loader';
import { frameAssetIds } from './frame';
import {
  PRESENTATION_SESSION_POLICIES,
  type PresentationSessionEvents,
  type PresentationSessionFont,
  type PresentationSessionMethods,
  type PresentationSlideSummary,
  type PresentationWireFrame,
} from './methods';

type Events = { [K in keyof PresentationSessionEvents]: PresentationSessionEvents[K] };

function slides(snapshot: DeckSnapshot): PresentationSlideSummary[] {
  return snapshot.slides.map((slide, index) => ({
    id: slide.id, index, name: slide.name, layoutPartPath: slide.layoutPartPath,
  }));
}

function size(snapshot: DeckSnapshot): { width: number; height: number } {
  return { width: snapshot.widthEmu, height: snapshot.heightEmu };
}

function fonts(faces?: readonly PresentationSessionFont[]): PptxFontFace[] | undefined {
  return faces?.map((face) => ({ ...face, bytes: new Uint8Array(face.bytes) }));
}

/** Creates a worker-owned presentation host with lazy wasm initialization. */
export function createPresentationSessionHost(
  transport: SessionTransport,
  options: { initWasm?: (source?: ArrayBuffer | WebAssembly.Module) => Promise<void> } = {}
): SessionHost<Events> {
  let handle: PresentationHandle | undefined;
  let disposed = false;
  let version = 0;
  let dirty = false;
  let epoch = 0;
  let slideCount: { version: string; count: number } | undefined;
  const encoder = new TextEncoder();
  const sentMedia = new Set<string>();

  function presentation(): PresentationHandle {
    if (disposed) throw new Error('Presentation session is disposed');
    if (!handle) throw new Error('Presentation session is not open');
    return handle;
  }

  function dispose(): void {
    disposed = true;
    const opened = handle;
    handle = undefined;
    sentMedia.clear();
    slideCount = undefined;
    opened?.dispose();
  }

  const handlers: MethodHandlers<PresentationSessionMethods, null> = {
    async open(_, bytes, input = {}) {
      if (disposed) throw new Error('Presentation session is disposed');
      if (handle) throw new Error('Presentation session is already open');
      await (options.initWasm ?? initWasm)(input.wasm);
      if (disposed) throw new Error('Presentation session is disposed');
      const opened = openPresentation(new Uint8Array(bytes), {
        clientId: input.clientId,
        fonts: fonts(input.fonts),
        fallbackFonts: fonts(input.fallbackFonts),
        initialUpdate: input.initialUpdate,
      });
      try {
        const snapshot = opened.snapshot();
        handle = opened;
        return { format: 'pptx', stage: 'ready', version, dirty,
          slides: slides(snapshot), size: size(snapshot) };
      } catch (error) {
        opened.dispose();
        throw error;
      }
    },
    version: () => presentation().version(),
    readContent: (_, request) => presentation().readContent(request),
    findText: (_, request) => presentation().findText(request),
    validateEdits: (_, request) => presentation().validateEdits(request),
    applyEdits(_, request) {
      const result = presentation().applyEdits(request);
      if (result.ok && result.applied) {
        version += 1;
        dirty = true;
        host.emit('changed', { version, dirty });
      }
      return result;
    },
    frame(_, slideIndex) {
      const opened = presentation();
      const frameVersion = opened.version();
      if (slideCount?.version !== frameVersion) {
        slideCount = { version: frameVersion, count: opened.snapshot().slides.length };
      }
      if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= slideCount.count) {
        throw new RangeError('Slide index is out of range');
      }
      const json = presentationDisplayListJson(opened, slideIndex);
      const displayList = encoder.encode(json).buffer;
      const media: PresentationWireFrame['media'] = [];
      for (const assetId of frameAssetIds(JSON.parse(json) as SlideDisplayList)) {
        if (sentMedia.has(assetId)) continue;
        let bytes: Uint8Array;
        try { bytes = opened.mediaBytes(assetId); } catch { continue; }
        if (isTiff(bytes) && bytes.byteLength <= MAX_TIFF_BYTES) {
          try { bytes = decodeTiffImage(bytes); } catch {}
        }
        media.push({ assetId, bytes: new Uint8Array(bytes).buffer });
      }
      epoch += 1;
      for (const { assetId } of media) sentMedia.add(assetId);
      return transferable({ displayList, version: frameVersion, epoch, slideIndex, media },
        [displayList, ...media.map(({ bytes }) => bytes)]);
    },
    slides: () => slides(presentation().snapshot()),
    slideSize: () => size(presentation().snapshot()),
    save() {
      const bytes = presentation().save();
      const buffer = bytes.buffer instanceof ArrayBuffer && bytes.byteOffset === 0 &&
        bytes.byteLength === bytes.buffer.byteLength ? bytes.buffer : new Uint8Array(bytes).buffer;
      return transferable(buffer, [buffer]);
    },
    dispose() {
      presentation();
      dispose();
    },
  };

  const host = createSessionHost<PresentationSessionMethods, Events, null>(transport, {
    handlers, policies: PRESENTATION_SESSION_POLICIES, context: null, onDispose: dispose,
  });
  return host;
}
