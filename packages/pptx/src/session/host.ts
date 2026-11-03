import { metafileImageBytes } from '../render/browserImage';
import { createSessionHost, type SessionHost } from '../../../../shared/office-session/host';
import { createCommandFontInstallation } from './commandFontInstallation';
import { isHostMessage, transferable, type TransferResult } from '../../../../shared/office-session/protocol';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import type { MethodHandlers } from '../../../../shared/office-session/types';
import { compileSessionWasm } from '#session-scope';
import { MAX_TIFF_BYTES, isTiff } from '../../../../shared/media';
import type { PptxFontFace, SlideDisplayList } from '../types';
import { wasmAssetUrl } from '../wasm/asset';
import { engineKind } from '../wasm/engine';
import {
  decodeTiffImage, initWasm, openPresentation, presentationDisplayListJson, presentationMetadata,
  openPresentationReplayBaseline, presentationPeerHydration, presentationPeerMetadata,
  presentationPeerDisplayListJson, registerPresentationPeerFonts, replayPresentation,
  type PresentationHandle, type PresentationPeerHandle,
} from '../wasm/loader';
import { frameAssetIds } from './frame';
import { createMetafileSvg } from './metafileSvg';
import { registerPresentationCommandHandlers } from './commandHost';
import type { listenPptxWorkerRender } from './workerRender';
import { PPTX_RENDER_REPLY, type PptxRenderReply } from '../../../../shared/office-session/render/pptxWorkerRender';
import type { PresentationCommandEvents, PresentationCommandMethods } from './commandProtocol';
import type { PresentationSharedMethods } from './commandSharedAdapter';
import type { DefaultFontInstallRequest, DefaultFontPayload } from '../../../../shared/pptxDefaultFonts';
import { peerHydrationError } from './editorPeerHydrationError';
import { PptxPeerHydrationError } from './peerHydrationError';
import {
  PRESENTATION_EDITOR_POLICIES, type PresentationEditorMethods,
} from './replay';
import {
  PRESENTATION_SESSION_POLICIES,
  type PresentationSessionEvents,
  type PresentationSessionFont,
  type PresentationSessionMethods,
  type PresentationWireFrame,
} from './methods';

type Events = { [K in keyof PresentationSessionEvents]: PresentationSessionEvents[K] } & PresentationCommandEvents & {
  defaultFonts: PresentationCommandEvents['ack'] & DefaultFontPayload;
  defaultFontInstall: DefaultFontInstallRequest;
};

function fonts(faces?: readonly PresentationSessionFont[]): PptxFontFace[] | undefined {
  return faces?.map((face) => ({ ...face, bytes: new Uint8Array(face.bytes) }));
}

export async function initializePresentationEditorWasm(source?: ArrayBuffer | WebAssembly.Module) {
  if (engineKind === 'native') {
    await initWasm(source);
    return undefined;
  }
  const module = source instanceof WebAssembly.Module ? source : source === undefined
    ? await compileSessionWasm(wasmAssetUrl()) : await WebAssembly.compile(source);
  await initWasm(module);
  return module;
}

/** Creates a worker-owned presentation host with lazy wasm initialization. */
export function createPresentationSessionHost(
  transport: SessionTransport,
  options: {
    workerRender?: ReturnType<typeof listenPptxWorkerRender>;
    initWasm?: (source?: ArrayBuffer | WebAssembly.Module) => Promise<void>;
    initEditorWasm?: (source?: ArrayBuffer | WebAssembly.Module) => Promise<WebAssembly.Module | undefined>;
  } = {}
): SessionHost<Events> {
  let handle: PresentationHandle | PresentationPeerHandle | undefined;
  let editor = false;
  let attached = false;
  let attachmentAttempted = false;
  let editorFailed = false;
  let sequence = 0;
  let disposed = false;
  let version = 0;
  let dirty = false;
  let epoch = 0;
  let metadataCache: { version: string; value: ReturnType<typeof presentationMetadata> } | undefined;
  const encoder = new TextEncoder();
  const sentMedia = new Set<string>();
  const metafiles = engineKind === 'native' ? {
    start: async () => {}, ready: () => true,
    convert: (assetId: string, _bytes: Uint8Array) => presentation().metafileSvg(assetId),
  } : createMetafileSvg();

  async function loadMetafiles(): Promise<void> {
    try {
      await metafiles.start();
    } catch {}
  }

  function presentation(): PresentationHandle | PresentationPeerHandle {
    if (disposed) throw new Error('Presentation session is disposed');
    if (!handle) throw new Error('Presentation session is not open');
    return handle;
  }

  function metadata(
    opened = presentation(), currentVersion = opened.version()
  ): ReturnType<typeof presentationMetadata> {
    if (metadataCache?.version !== currentVersion) {
      metadataCache = { version: currentVersion, value: editor
        ? presentationPeerMetadata(opened) : presentationMetadata(opened as PresentationHandle) };
    }
    return metadataCache.value;
  }

  function editorPresentation(requireAttachment = true): PresentationPeerHandle {
    const opened = presentation();
    if (!editor || editorFailed || (requireAttachment && !attached)) {
      throw new PptxPeerHydrationError('stage', 'Presentation editor is not attached');
    }
    return opened;
  }

  function position(expectedSequence: number, expectedVersion: string) {
    const opened = editorPresentation();
    if (sequence !== expectedSequence || opened.version() !== expectedVersion) {
      editorFailed = true;
      throw new PptxPeerHydrationError('position-mismatch', 'Presentation editor fence differs');
    }
    return { sequence, version: opened.version() };
  }

  function projection() {
    const value = metadata();
    return { format: 'pptx' as const, stage: 'ready' as const, version, dirty,
      slides: value.slides, size: value.size };
  }

  function dispose(): void {
    disposed = true;
    const opened = handle;
    handle = undefined;
    sentMedia.clear();
    metadataCache = undefined;
    opened?.dispose();
  }

  const handlers: MethodHandlers<PresentationSessionMethods, null> = {
    async open(_, bytes, input = {}) {
      if (disposed) throw new Error('Presentation session is disposed');
      if (handle || editor) throw new Error('Presentation session is already open');
      await (options.initWasm ?? initWasm)(input.wasm);
      if (disposed) throw new Error('Presentation session is disposed');
      const opened = openPresentation(new Uint8Array(bytes), {
        clientId: input.clientId,
        fonts: fonts(input.fonts),
        fallbackFonts: fonts(input.fallbackFonts),
        initialUpdate: input.initialUpdate, retainSource: false,
      });
      try {
        const projection = metadata(opened);
        handle = opened;
        return { format: 'pptx', stage: 'ready', version, dirty,
          slides: projection.slides, size: projection.size };
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
      if (editor) throw new PptxPeerHydrationError('stage', 'Editor mutations require ordered replay');
      const result = (presentation() as PresentationHandle).applyEdits(request);
      if (result.ok && result.applied) {
        version += 1;
        dirty = true;
        host.emit('changed', { version, dirty });
      }
      return result;
    },
    async frame(_, slideIndex) {
      const opened = presentation();
      const frameVersion = opened.version();
      const slideCount = metadata(opened, frameVersion).slides.length;
      if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= slideCount) {
        throw new RangeError('Slide index is out of range');
      }
      const json = editor ? presentationPeerDisplayListJson(opened, slideIndex)
        : presentationDisplayListJson(opened as PresentationHandle, slideIndex);
      const displayList = encoder.encode(json).buffer;
      const media: PresentationWireFrame['media'] = [];
      const sources = new Map<string, Uint8Array>();
      let needsMetafiles = false;
      for (const assetId of frameAssetIds(JSON.parse(json) as SlideDisplayList)) {
        if (sentMedia.has(assetId)) continue;
        let bytes: Uint8Array;
        try { bytes = opened.mediaBytes(assetId); } catch { continue; }
        if (isTiff(bytes) && bytes.byteLength <= MAX_TIFF_BYTES) {
          try { bytes = decodeTiffImage(bytes); } catch {}
        }
        sources.set(assetId, bytes);
        if (!metafiles.ready() && !needsMetafiles) {
          metafileImageBytes(bytes, () => { needsMetafiles = true; return undefined; });
        }
      }
      if (needsMetafiles) {
        await loadMetafiles();
        presentation();
      }
      const pendingMetafiles = new Set<string>();
      for (const [assetId, source] of sources) {
        const bytes = metafileImageBytes(source, () => {
          if (metafiles.ready()) return metafiles.convert(assetId, source);
          pendingMetafiles.add(assetId);
          return undefined;
        });
        media.push({ assetId, bytes: new Uint8Array(bytes).buffer });
      }
      epoch += 1;
      for (const { assetId } of media) if (!pendingMetafiles.has(assetId)) sentMedia.add(assetId);
      return transferable({ displayList, version: frameVersion, epoch, slideIndex, media },
        [displayList, ...media.map(({ bytes }) => bytes)]);
    },
    slides: () => metadata().slides,
    slideSize: () => metadata().size,
    save() {
      if (editor) throw new PptxPeerHydrationError('stage', 'Editor saves require a fence');
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

  const editorHandlers: MethodHandlers<PresentationEditorMethods, null> = {
    async beginEditor(_, bytes, input) {
      if (disposed) throw new Error('Presentation session is disposed');
      if (handle || editor) throw new Error('Presentation session is already open');
      editor = true;
      const module = await (options.initEditorWasm ?? initializePresentationEditorWasm)(input.wasm);
      if (disposed) throw new Error('Presentation session is disposed');
      let opened: PresentationPeerHandle | undefined;
      try {
        opened = openPresentationReplayBaseline(new Uint8Array(bytes), {
          clientId: input.clientId, fonts: fonts(input.fonts), fallbackFonts: fonts(input.fallbackFonts),
          initialUpdate: input.initialUpdate,
        });
        await registerPresentationPeerFonts(opened);
        if (disposed) throw new Error('Presentation session is disposed');
        const hydration = presentationPeerHydration(opened);
        handle = opened;
        return { state: projection(), hydration, module, sequence, version: opened.version() };
      } catch (error) {
        if (handle === opened) handle = undefined;
        opened?.dispose();
        editorFailed = true;
        throw peerHydrationError(error);
      }
    },
    attachPeer(_, peerVersion, peerSequence) {
      const alreadyAttempted = attachmentAttempted;
      attachmentAttempted = true;
      attached = false;
      const opened = editorPresentation(false);
      if (alreadyAttempted || peerVersion !== opened.version() || peerSequence !== sequence) {
        editorFailed = true;
        throw new PptxPeerHydrationError(alreadyAttempted ? 'attachment' : 'position-mismatch',
          'Presentation peer attachment differs or was already attempted');
      }
      attached = true;
      return { version: opened.version(), sequence };
    },
    replay(_, envelope) {
      const opened = editorPresentation();
      try {
        if (envelope.expectedOutcome == null) {
          throw new PptxPeerHydrationError('outcome', 'Editor replay requires an expected outcome');
        }
        const reply = replayPresentation(opened, envelope);
        if (!reply.consumed) throw new PptxPeerHydrationError('refusal', 'Worker refused accepted input');
        sequence = reply.sequence;
        version = reply.revision;
        dirty ||= reply.outcome.applied;
        return { ...reply, projection: projection() };
      } catch (error) {
        editorFailed = true;
        throw peerHydrationError(error);
      }
    },
    flush: (_, expectedSequence, expectedVersion) => position(expectedSequence, expectedVersion),
    async editorFrame(_, slideId) {
      editorPresentation(false);
      const index = metadata().slides.findIndex((slide) => slide.id === slideId);
      const frameSequence = sequence;
      const frame = await handlers.frame(null, index) as TransferResult<PresentationWireFrame>;
      return transferable({ ...frame.value, slideId, sequence: frameSequence }, frame.transfer);
    },
    editorSave(_, expectedSequence, expectedVersion) {
      position(expectedSequence, expectedVersion);
      const buffer = new Uint8Array(editorPresentation().save()).buffer;
      return transferable(buffer, [buffer]);
    },
  };

  let pipeOutputTimer: ReturnType<typeof setTimeout> | undefined;
  const fontInstallation = createCommandFontInstallation((request, transfer) => host.emit('defaultFontInstall', request, transfer));
  const commands = registerPresentationCommandHandlers({ initWasm: options.initWasm,
    onActiveFrame: options.workerRender?.frame,
    activeFrameWanted: options.workerRender?.active,
    onResetMedia: options.workerRender?.resetMedia,
    installDefaultFonts: async (faces) => {
      const installed = await fontInstallation.install(faces);
      options.workerRender?.defaultFonts(installed);
      return installed;
    },
    onDefaultFonts: (result) => host.emit('defaultFonts', result.value, result.transfer), onPipeOutput() {
    if (disposed || pipeOutputTimer !== undefined) return;
    pipeOutputTimer = setTimeout(() => {
      pipeOutputTimer = undefined;
      const output = commands.drainPipe();
      if (output) host.emit('collaborationPipe', output.batches, output.transfer);
    }, 0);
  } });
  let queued: Promise<void> | undefined;
  const send = (message: unknown, transfer: Transferable[] | undefined, reply?: PptxRenderReply) => {
    if (reply && !disposed) {
      try {
        transport.post({ ...(message as object), [PPTX_RENDER_REPLY]: reply },
          reply.kind === 'raster' ? [...(transfer ?? []), reply.bitmap] : transfer);
        return;
      } catch {}
    }
    if (reply?.kind === 'raster') reply.bitmap.close();
    transport.post(message, transfer);
  };
  const replyTransport: SessionTransport = { ...transport,
    listen: (listener) => transport.listen((message) => { if (!fontInstallation.receive(message)) listener(message); }),
    post(message, transfer) {
      const output = isHostMessage(message) && (message.kind === 'reply' || message.kind === 'event' && message.name === 'ack')
        ? commands.drainPipe() : undefined;
      const value = output ? { ...(message as object), collabPipe: output.batches } : message;
      const list = output ? [...(transfer ?? []), ...output.transfer] : transfer;
      const render = options.workerRender;
      if (!queued && !render?.pending()) {
        transport.post(value, list);
        render?.afterReply();
        return;
      }
      const reply = render?.pending() ? render.paintNow() : undefined;
      const next = Promise.all([queued, reply]).then(([, painted]) => {
        try { send(value, list, painted); } catch (error) {
          try {
            transport.post({ protocol: 1, kind: 'failure', code: 'message',
              message: `Session could not post a message: ${error instanceof Error ? error.message : String(error)}` });
          } catch {}
        }
        render?.afterReply();
      });
      queued = next;
      void next.then(() => { if (queued === next) queued = undefined; });
    },
  };
  const host = createSessionHost<
    PresentationSessionMethods & PresentationEditorMethods & PresentationCommandMethods & PresentationSharedMethods, Events, null
  >(replyTransport, {
    handlers: { ...handlers, ...editorHandlers, ...commands.handlers },
    policies: { ...PRESENTATION_SESSION_POLICIES, ...PRESENTATION_EDITOR_POLICIES, ...commands.policies },
    context: null, onDispose() {
      if (pipeOutputTimer !== undefined) clearTimeout(pipeOutputTimer);
      try { dispose(); } finally { fontInstallation.dispose(); commands.dispose(); options.workerRender?.dispose(); }
    },
  });
  return host;
}
