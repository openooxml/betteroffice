import { createPptxWorkerSession, presentationSharedInputs, PRESENTATION_COMMAND_ADMISSION } from '@betteroffice/pptx/worker';
import { sizeCanvasForSlide, PptxWorkerEditorDisposedError, PptxWorkerEditorFailedError } from '@betteroffice/pptx/editor';
import type {
  DeckSnapshot, PptxFontFace, PptxWorkerSession, PresentationCommandClient, PresentationCommandFrame, PresentationCommandAck,
} from '@betteroffice/pptx';
import { useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { RefObject } from 'react';
import { flushSync } from 'react-dom';
import type { PptxAsyncEditorProps } from '../PptxEditor';
import type { SlideFit } from '../fitSlideScale';
import { installFonts, useStableFonts } from './peerlessFonts';
import { createCommandFrameImages, createCommandFramePainter, type CommandFrameImages } from './commandFramePainter';
import { createLocalEchoPainter } from './localEchoPainter';
import {
  createPeerlessViewStore, installCoveringFrame,
  type PeerlessConfirmedFrame, type PeerlessViewSnapshot, type PeerlessViewStore,
} from './peerlessViewStore';

export const peerlessSessionOpener = { open: createPptxWorkerSession };
export const peerlessViewStoreFactory = {
  create: (...args: Parameters<typeof createPeerlessViewStore>) => createPeerlessViewStore(...args),
};
export const peerlessPainterFactories = {
  frame: createCommandFramePainter,
  echo: createLocalEchoPainter,
};

function idle(callback: () => void): () => void {
  if (typeof window.requestIdleCallback === 'function') {
    const id = window.requestIdleCallback(callback);
    return () => window.cancelIdleCallback(id);
  }
  const id = window.setTimeout(callback, 0);
  return () => window.clearTimeout(id);
}

export type PeerlessVisualStore = Pick<PeerlessViewStore, 'getSnapshot' | 'subscribePresentation'>;
export interface PeerlessVisualSnapshot { view: PeerlessViewSnapshot; fit: SlideFit }

export interface PeerlessPresentation {
  readonly fileGeneration: object;
  readonly session: PptxWorkerSession;
  readonly client: PresentationCommandClient;
  readonly store: PeerlessViewStore;
  readonly images: CommandFrameImages;
  readonly visual: PeerlessVisualStore;
  readonly ready: Promise<void>;
  readonly current: boolean;
  readonly selectionEnabled: boolean;
  readonly inputEnabled: boolean;
  readonly hydrated: boolean;
  readonly failure: PptxWorkerEditorFailedError | undefined;
  getVisualSnapshot(): PeerlessVisualSnapshot | null;
  subscribeVisual(listener: () => void): () => void;
  require(): void;
  fit(value: SlideFit): void;
  installFont(face: PptxFontFace): Promise<void>;
  setReadOnly(readOnly: boolean): void;
  setChangeObserver(enabled: boolean): void;
  recover(): Promise<{ bytes: Uint8Array; recovery: true; throughSeq: number; version: string }>;
}

export function usePeerlessPresentation(
  props: PptxAsyncEditorProps, canvas: RefObject<HTMLCanvasElement | null>,
): { presentation: PeerlessPresentation | null; pending: PeerlessPresentation | null; failure: Error | null } {
  const fonts = useStableFonts(props.fonts);
  const collaboration = props.collaboration;
  const initialUpdate = useRef(collaboration?.initialUpdate);
  const update = collaboration?.initialUpdate;
  if (initialUpdate.current !== update && (!initialUpdate.current || !update ||
    initialUpdate.current.length !== update.length || !initialUpdate.current.every((byte, index) => byte === update[index]))) {
    initialUpdate.current = update;
  }
  const fileGeneration = useMemo(() => ({}), [props.file, fonts, props.clientId, Boolean(collaboration),
    collaboration?.roomId, collaboration?.clientId, initialUpdate.current]);
  const latestFile = useRef(fileGeneration);
  latestFile.current = fileGeneration;
  const callbacks = useRef(props);
  callbacks.current = props;
  const [presentation, setPresentation] = useState<PeerlessPresentation | null>(null);
  const [pending, setPending] = useState<PeerlessPresentation | null>(null);
  const requestMode = useRef<((readOnly: boolean) => void) | null>(null);
  const [failure, setFailure] = useState<Error | null>(null);
  const mode = useRef(props.readOnly ?? false);
  useLayoutEffect(() => {
    mode.current = props.readOnly ?? false;
    requestMode.current?.(mode.current);
  }, [props.readOnly]);

  useLayoutEffect(() => {
    setPresentation(null);
    setPending(null);
    setFailure(null);
    if (!props.file || !canvas.current) return;
    const surface = canvas.current;
    let alive = true;
    let session: PptxWorkerSession | undefined;
    let observedChanges = Boolean(callbacks.current.onChange);
    let base: PeerlessViewStore | undefined;
    let run: PeerlessPresentation | undefined;
    let framePainter: ReturnType<typeof createCommandFramePainter> | undefined;
    let displayedPainter: ReturnType<typeof createCommandFramePainter> | undefined;
    let echoPainter: ReturnType<typeof createLocalEchoPainter> | undefined;
    const images = createCommandFrameImages();
    let activeImages: ReturnType<typeof images.pin> | undefined;
    let fatal: PptxWorkerEditorFailedError | undefined;
    let failedView: PeerlessViewSnapshot | undefined;
    let fit: SlideFit | undefined;
    let queuedFrame: PeerlessConfirmedFrame | undefined;
    let visual: PeerlessVisualSnapshot | null = null;
    let cancelIdle: (() => void) | undefined;
    let draining = false;
    let changeSeq: number | undefined;
    let changeEvent: PresentationCommandAck['change'];
    let checkpointTimer: ReturnType<typeof setTimeout> | undefined;
    const recoveryBase = props.file instanceof Uint8Array ? props.file.slice() : new Uint8Array(props.file).slice();
    let checkpoint = { bytes: recoveryBase, throughSeq: 0, version: '' };
    let checkpointUpdate: Uint8Array | undefined;
    let journal: { throughSeq: number; version: string; updates: Uint8Array[] }[] = [];
    let firstPainted = false;
    let published = false;
    let paintedFrame: PresentationCommandFrame | undefined;
    let paintedView: PeerlessViewSnapshot | undefined;
    let modeSequence = 0;
    let acknowledgedMode = 0;
    let requestedMode = mode.current;
    let modeClient: PresentationCommandClient | undefined;
    const queuedModes: { sequence: number; readOnly: boolean }[] = [];
    const openingModes = new Set<Promise<unknown>>();
    const dirty = new Map<string, number>();
    const vanished = new Set<string>();
    const pendingMedia = new Map<string, ArrayBuffer>();
    const installedFonts: FontFace[] = [];
    const disposers: (() => void)[] = [];
    const failureListeners = new Set<() => void>();
    const visualListeners = new Set<() => void>();
    let wake!: () => void;
    let changed = new Promise<void>((resolve) => { wake = resolve; });
    const signal = () => { wake(); changed = new Promise<void>((resolve) => { wake = resolve; }); };
    const currentFile = () => alive && latestFile.current === fileGeneration;
    const current = () => currentFile() && !fatal;
    const require = () => {
      if (!currentFile()) throw new PptxWorkerEditorDisposedError();
      if (fatal) throw fatal;
    };
    const modeReady = () => acknowledgedMode === modeSequence;
    const hydrated = () => current() && modeReady() && Boolean(base) && (base!.getSnapshot().confirmed === null ||
      firstPainted && Boolean(fit?.scale) && visual?.fit === fit && paintedView?.confirmed === queuedFrame);
    const selectionEnabled = () => current() && published && Boolean(base?.getSnapshot().confirmed);
    const inputEnabled = () => selectionEnabled() && modeReady() && !mode.current &&
      !base?.getSnapshot().editState.readOnly;
    const notifyVisual = () => { for (const listener of visualListeners) listener(); };
    const report = (cause: unknown) => {
      if (!current()) return;
      fatal = cause instanceof PptxWorkerEditorFailedError ? cause : new PptxWorkerEditorFailedError(cause);
      const event = changeEvent;
      if (changeSeq !== undefined && event && modeClient && event.throughSeq >= changeSeq &&
        event.throughSeq <= modeClient.state.throughSeq && event.version === modeClient.state.version) {
        changeSeq = undefined;
        changeEvent = undefined;
        try { callbacks.current.onChange?.(event.snapshot); } catch {}
      }
      cancelIdle?.(); cancelIdle = undefined;
      clearTimeout(checkpointTimer); checkpointTimer = undefined;
      framePainter?.freeze();
      if (framePainter !== displayedPainter) framePainter?.dispose();
      echoPainter?.dispose(); echoPainter = undefined;
      activeImages?.release(); activeImages = undefined;
      images.dispose();
      for (const face of installedFonts.splice(0)) document.fonts.delete(face);
      if (base) {
        const view = paintedView ?? base.getSnapshot();
        failedView = { ...view, focused: false, caret: null, selectionRects: [],
          composition: null, echo: null, pending: [], waitingInputCount: 0,
          editState: { ...view.editState, readOnly: true,
            toolbar: { ...view.editState.toolbar, canEditText: false } } };
        if (visual) visual = { ...visual, view: failedView };
      }
      signal(); notifyVisual();
      for (const listener of failureListeners) { try { listener(); } catch {} }
      setFailure(fatal);
      try { callbacks.current.onError?.(fatal); } catch {}
      void session?.dispose().catch(() => {});
    };
    const fontsReady = installFonts(fonts, installedFonts, current);
    void fontsReady.catch(report);

    async function whenPainted(): Promise<void> {
      while (true) {
        require();
        const next = changed;
        if (hydrated()) return;
        await next;
      }
    }

    async function waitForMode(token: SlideFit, currentPaint: () => boolean): Promise<void> {
      while (current() && fit === token && currentPaint() && !modeReady()) await changed;
    }

    function paintLocal(view: PeerlessViewSnapshot, restore = false): void {
      if (!current() || !firstPainted || !visual || fit !== visual.fit || view.confirmed !== visual.view.confirmed) return;
      if (restore) echoPainter?.restoreLine(view); else echoPainter?.paintEcho(view);
      paintedView = view;
      visual = { ...visual, view };
      notifyVisual();
    }

    function paintFrame(frame: PeerlessConfirmedFrame, view: PeerlessViewSnapshot): void {
      if (!current()) return;
      const nextImages = images.pin(frame, [...pendingMedia].map(([assetId, bytes]) => ({ assetId, bytes })));
      activeImages?.release(); activeImages = nextImages;
      queuedFrame = frame;
      firstPainted = false;
      cancelIdle?.(); cancelIdle = undefined;
      signal();
      if (!fit?.scale || !framePainter) return;
      const supplied = [...pendingMedia].map(([assetId, bytes]) => ({ assetId, bytes }));
      pendingMedia.clear();
      try { framePainter.paintFrame(frame, view, supplied); }
      catch (error) { if (modeClient && !presentationSharedInputs.get(modeClient)) report(error); throw error; }
      void framePainter.whenPainted().catch(report);
    }

    function sendModes(): void {
      const client = modeClient;
      if (!current() || !client) return;
      for (const { sequence, readOnly } of queuedModes.splice(0)) {
        const operation: Promise<void> = (base ?? client).dispatch({ kind: 'setReadOnly', payload: { readOnly,
          inputTime: performance.now(), undoGroup: 'mode' } }).then((ack) => {
          if (!current() || sequence !== modeSequence || ack.editState.readOnly !== requestedMode) return;
          acknowledgedMode = sequence;
          for (const listener of failureListeners) listener();
          schedule(client);
        }).catch(report).finally(() => { openingModes.delete(operation); signal(); });
        if (!base) openingModes.add(operation);
      }
    }

    function setReadOnly(readOnly: boolean): void {
      if (!current() || requestedMode === readOnly) return;
      requestedMode = readOnly;
      queuedModes.push({ sequence: ++modeSequence, readOnly });
      signal();
      for (const listener of failureListeners) listener();
      sendModes();
    }
    requestMode.current = setReadOnly;

    function schedule(client: PresentationCommandClient): void {
      if (!current() || !base || (base.getSnapshot().confirmed && !firstPainted) || cancelIdle || draining) return;
      cancelIdle = idle(() => {
        cancelIdle = undefined;
        if (!current() || !base) return;
        const view = base.getSnapshot();
        if (view.confirmed && (!firstPainted || !paintedFrame ||
          paintedView?.confirmed !== view.confirmed || view.presentedThroughSeq < client.state.throughSeq)) return;
        draining = true;
        void drain(client).catch(report).finally(() => {
          draining = false;
          if (current() && (dirty.size > 0 || changeSeq !== undefined)) schedule(client);
        });
      });
    }

    async function saveCheckpoint(client: PresentationCommandClient): Promise<void> {
      await base?.fence(); await whenPainted();
      if (!current()) return;
      const result = await client.checkpoint?.();
      if (current() && result && result.throughSeq >= checkpoint.throughSeq) {
        checkpoint = { bytes: new Uint8Array(result.bytes).slice(), throughSeq: result.throughSeq, version: result.version };
        checkpointUpdate = result.recoveryUpdate ? new Uint8Array(result.recoveryUpdate).slice() : undefined;
        journal = journal.filter((entry) => entry.throughSeq > result.throughSeq);
      }
    }

    function scheduleCheckpoint(client: PresentationCommandClient): void {
      if (checkpointTimer) return;
      checkpointTimer = setTimeout(() => { checkpointTimer = undefined; void saveCheckpoint(client).catch(report); }, 1000);
    }

    async function drain(client: PresentationCommandClient): Promise<void> {
      if (!base || !current()) return;
      await whenPainted();
      const pendingSeq = changeSeq;
      if (pendingSeq !== undefined) {
        if (callbacks.current.onChange) {
          await base.fence();
          if (!current()) return;
          await whenPainted();
          const event = changeEvent;
          if (event) {
            changeEvent = undefined;
            if (changeSeq !== undefined && changeSeq <= event.throughSeq) changeSeq = undefined;
            if (event.throughSeq <= client.state.throughSeq && event.version === client.state.version) {
              try { callbacks.current.onChange?.(event.snapshot); } catch {}
            }
          } else {
          const reply = await client.read();
          if (!current()) return;
          await whenPainted();
          let snapshot: DeckSnapshot | undefined;
          if (reply.result.ok) {
            snapshot = client.access ? await client.access('snapshot') : { widthEmu: client.state.size.width,
              heightEmu: client.state.size.height, slides: reply.result.slides };
            if (!current()) return;
          }
          if (changeSeq !== undefined && changeSeq <= Math.max(pendingSeq, reply.throughSeq)) changeSeq = undefined;
          if (snapshot) { try { callbacks.current.onChange?.(snapshot); } catch {} }
          }
        } else changeSeq = undefined;
      }
      if (!current()) return;
      await whenPainted();
      const activeId = base.getSnapshot().confirmed?.frame.slideId;
      if (activeId) dirty.delete(activeId);
      const entry = dirty.entries().next().value;
      if (!entry) return;
      const [slideId, generation] = entry;
      const cached = base.getSnapshot().thumbnails.get(slideId);
      const reply = await client.thumbnail(slideId, cached?.frame.slideRevision).catch((error: unknown) => {
        if (!client.state.slides.some((slide) => slide.id === slideId)) return null;
        throw error;
      });
      if (!reply) return;
      if (!current()) return;
      if (reply.session !== client.state.session || reply.incarnation !== client.state.incarnation ||
        reply.epoch !== client.state.epoch) return;
      if (reply.status === 'refused') { vanished.add(slideId); dirty.delete(slideId); return; }
      if (reply.status === 'frame') {
        for (const asset of reply.frame.media) {
          pendingMedia.set(asset.assetId, asset.bytes);
        }
        base.installMedia(reply.frame.media);
      }
      await whenPainted();
      if (dirty.get(slideId) !== generation) return;
      if (reply.status === 'frame') {
        base.installThumbnail(reply.frame);
      }
      if (dirty.get(slideId) === generation) dirty.delete(slideId);
    }

    try {
      session = peerlessSessionOpener.open(props.file, {
        fonts, clientId: props.clientId, collaboration: props.collaboration,
        readOnly: mode.current, observeChanges: observedChanges,
        activeSlideIndex: Number.isInteger(props.initialSlide) ? Math.max(0, props.initialSlide! - 1) : 0,
      });
      disposers.push(session.subscribeFailure(report));
      const owner = session;
      void owner.ready.then((client) => {
        if (!current()) return;
        const cacheMedia = (assets: PresentationCommandFrame['media']) => {
          if (!current()) return;
          for (const asset of assets) pendingMedia.set(asset.assetId, asset.bytes);
        };
        disposers.push(client.subscribeMedia(cacheMedia));
        modeClient = client;
        if (!queuedModes.length) {
          requestedMode = client.state.editState.readOnly;
          setReadOnly(mode.current);
        }
        sendModes();
      }).catch(report);
      void Promise.all([owner.ready, fontsReady]).then(async ([client]) => {
        while (current() && openingModes.size) await Promise.all([...openingModes]);
        if (!current()) return;
        checkpoint.version = client.state.version;
        checkpointUpdate = client.state.recoveryUpdate ? new Uint8Array(client.state.recoveryUpdate).slice() : undefined;
        base = peerlessViewStoreFactory.create(client, {
          checkpoint(pixels) {
            const restoreFrame = framePainter?.checkpoint(pixels);
            const restoreEcho = echoPainter?.checkpoint();
            const savedImages = activeImages?.retain();
            const saved = { queuedFrame, paintedFrame, paintedView, firstPainted, fit, visual,
              pendingMedia: new Map(pendingMedia) };
            let settled = false;
            const finish = () => {
              if (settled) return;
              settled = true;
              restoreFrame?.commit?.(); restoreEcho?.commit?.(); savedImages?.release();
            };
            return Object.assign(() => {
              if (settled) return;
              if (current() && fit === saved.fit) {
                ({ queuedFrame, paintedFrame, paintedView, firstPainted, visual } = saved);
                activeImages?.release(); activeImages = savedImages?.retain();
                pendingMedia.clear();
                for (const [assetId, bytes] of saved.pendingMedia) pendingMedia.set(assetId, bytes);
              }
              try {
                restoreFrame?.(); restoreEcho?.();
                void framePainter?.whenPainted().catch(report);
                signal(); schedule(client);
              } finally { finish(); }
            }, { commit: finish });
          },
          paintFrame,
          paintEcho: (view) => paintLocal(view),
          restoreLine: (view) => paintLocal(view, true),
        }, { onError: (error) => callbacks.current.onError?.(new PptxWorkerEditorFailedError(error)) });
        const original = base;
        const getSnapshot = original.getSnapshot.bind(original);
        const subscribePresentation = original.subscribePresentation.bind(original);
        const subscribeEditState = original.subscribeEditState.bind(original);
        const dispatch = original.dispatch.bind(original);
        const fence = original.fence.bind(original);
        const setSelection = original.setSelection.bind(original);
        const setFocused = original.setFocused.bind(original);
        const beginComposition = original.beginComposition.bind(original);
        const updateComposition = original.updateComposition.bind(original);
        const installThumbnail = original.installThumbnail.bind(original);
        const subscribe = (listen: (listener: () => void) => () => void, listener: () => void) => {
          const off = listen(listener);
          failureListeners.add(listener);
          return () => { off(); failureListeners.delete(listener); };
        };
        let gatedSource: PeerlessViewSnapshot | undefined;
        let gatedView: PeerlessViewSnapshot | undefined;
        const guarded: Pick<PeerlessViewStore, 'getSnapshot' | 'subscribePresentation' | 'subscribeEditState' |
          'dispatch' | 'fence' | 'setSelection' | 'setFocused' | 'beginComposition' | 'updateComposition' | 'installThumbnail'> = {
          getSnapshot: () => {
            if (failedView) return failedView;
            const view = getSnapshot();
            if (!mode.current && modeReady()) return view;
            if (gatedSource !== view) {
              gatedSource = view;
              gatedView = { ...view, editState: { ...view.editState, readOnly: true } };
            }
            return gatedView!;
          },
          subscribePresentation: (listener) => subscribe(subscribePresentation, listener),
          subscribeEditState: (listener) => subscribe(subscribeEditState, listener),
          dispatch: (intent, options) => {
            try {
              require();
              const admission = PRESENTATION_COMMAND_ADMISSION[intent.kind];
              if (admission === 'local-write' && (mode.current || !modeReady() || getSnapshot().editState.readOnly)) {
                throw new Error('The editor is read-only');
              }
              if (admission !== 'mode' && intent.kind !== 'collabPipe' && !selectionEnabled() && !['undo', 'redo', 'applyEdits', 'registerFont', 'invoke', 'setChangeObserver'].includes(intent.kind)) {
                throw new Error('Slide input is unavailable');
              }
              return dispatch(intent, options);
            } catch (error) { return Promise.reject(error); }
          },
          fence: async () => { require(); await fence(); require(); },
          setSelection: (selection) => { require(); setSelection(selection); },
          setFocused: (focused) => { if (current()) setFocused(focused); },
          beginComposition: (id, time) => { require(); if (inputEnabled()) beginComposition(id, time); },
          updateComposition: (text) => { if (inputEnabled()) updateComposition(text); },
          installThumbnail: (frame) => current() && installThumbnail(frame),
        };
        const store = Object.assign(original, guarded);
        if (client.state.frame) installCoveringFrame(original, client.state.frame);
        for (const slide of client.state.slides) {
          if (slide.id !== client.state.activeSlideId) dirty.set(slide.id, 0);
        }
        disposers.push(client.subscribe((ack) => {
          if (!current()) return;
          if (ack.throughSeq > checkpoint.throughSeq && ack.updates?.length) {
            journal.push({ throughSeq: ack.throughSeq, version: ack.version,
              updates: ack.updates.map((update) => new Uint8Array(update.bytes).slice()) });
          }
          if (ack.slides) {
            const restored = new Set(ack.slides.map((slide) => slide.id));
            for (const id of vanished) if (restored.has(id)) vanished.delete(id);
          }
          for (const id of ack.dirtySlideIds) if (!vanished.has(id)) dirty.set(id, (dirty.get(id) ?? 0) + 1);
          if (ack.slides) {
            const ids = new Set(ack.slides.map((slide) => slide.id));
            for (const id of dirty.keys()) if (!ids.has(id)) dirty.delete(id);
            if (ack.activeSlideId === null) {
              queuedFrame = undefined;
              paintedFrame = undefined;
              paintedView = getSnapshot();
              visual = null;
              firstPainted = false;
              framePainter?.freeze();
              notifyVisual(); signal();
            }
          }
          for (const id of base!.getSnapshot().dirtyThumbnailIds) {
            if (!vanished.has(id) && !dirty.has(id)) dirty.set(id, 0);
          }
          if (ack.outcomes.some((outcome) => outcome.committed)) scheduleCheckpoint(client);
          if (callbacks.current.onChange && ack.outcomes.some((outcome) => outcome.committed)) { changeSeq = ack.throughSeq; changeEvent = ack.change; }
          schedule(client);
        }));
        const subscribeVisual = (listener: () => void) => {
          visualListeners.add(listener);
          return () => { visualListeners.delete(listener); };
        };
        run = {
          fileGeneration, session: owner, client, store, images,
          visual: { subscribePresentation: subscribeVisual, getSnapshot: () => failedView ?? paintedView ?? getSnapshot() },
          getVisualSnapshot: () => visual, subscribeVisual,
          get ready() { return whenPainted(); },
          get current() { return current(); },
          get selectionEnabled() { return selectionEnabled(); },
          get inputEnabled() { return inputEnabled(); },
          get hydrated() { return hydrated(); },
          get failure() { return fatal; },
          require, setReadOnly,
          setChangeObserver(enabled) {
            if (!current() || observedChanges === enabled) return;
            observedChanges = enabled;
            void store.dispatch({ kind: 'setChangeObserver', payload: { enabled, inputTime: performance.now(), undoGroup: 'observer' } }).catch(report);
          },
          async recover() {
            if (!currentFile()) throw new PptxWorkerEditorDisposedError();
            if (!fatal) await saveCheckpoint(client);
            if (!currentFile()) throw new PptxWorkerEditorDisposedError();
            if (!journal.length) return { ...checkpoint, bytes: checkpoint.bytes.slice(), recovery: true };
            if (!checkpointUpdate) throw new Error('Recovery baseline is unavailable');
            const recovering = peerlessSessionOpener.open(recoveryBase, { fonts, initialUpdate: checkpointUpdate });
            try {
              const replay = await recovering.ready;
              for (const entry of journal) for (const update of entry.updates) {
                if (!currentFile()) throw new PptxWorkerEditorDisposedError();
                const ack = await replay.dispatch({ kind: 'applyRemoteUpdate', payload: { bytes: update.slice().buffer,
                  inputTime: 0, undoGroup: 'recovery' } });
                if (ack.outcomes.some((outcome) => outcome.status === 'refused' || outcome.status === 'cancelled')) {
                  throw new Error('Recovery update refused');
                }
              }
              const bytes = await replay.save();
              if (!currentFile()) throw new PptxWorkerEditorDisposedError();
              const latest = journal[journal.length - 1]!;
              return { bytes, throughSeq: latest.throughSeq, version: latest.version, recovery: true };
            } finally { await recovering.dispose(); }
          },
          async installFont(face) {
            require();
            await installFonts([face], installedFonts, current);
            require();
          },
          fit(token) {
            if (!current() || !getSnapshot().confirmed || fit === token) return;
            fit = token;
            firstPainted = false;
            echoPainter?.dispose(); echoPainter = undefined;
            framePainter?.freeze();
            if (framePainter !== displayedPainter) framePainter?.dispose();
            framePainter = undefined;
            signal();
            if (!token.scale) return;
            let preparedEcho: ReturnType<typeof createLocalEchoPainter> | undefined;
            framePainter = peerlessPainterFactories.frame(surface, {
              ...token, images, current: (frame) => current() && fit === token && (!frame || frame === queuedFrame),
              beforeCommit: (currentPaint) => waitForMode(token, currentPaint),
              compose(raster, frame) {
                if (!current() || fit !== token || getSnapshot().confirmed !== frame) return;
                const view = getSnapshot();
                if (!view.echo && view.selection?.anchor === view.selection?.focus) return;
                const pixels = surface.ownerDocument.createElement('canvas');
                sizeCanvasForSlide(pixels, frame.displayList, token.dpr, token.scale);
                const ctx = pixels.getContext('2d');
                if (!ctx) { pixels.width = pixels.height = 0; throw new Error('Slide canvas is unavailable'); }
                let preparingEcho = true;
                let echo: ReturnType<typeof createLocalEchoPainter> | undefined;
                try {
                  ctx.drawImage(raster.canvas, 0, 0);
                  echo = peerlessPainterFactories.echo(pixels, { ...token,
                    current: () => current() && fit === token && (preparingEcho || firstPainted) });
                  echo.paintEcho(view);
                } catch (error) { echo?.dispose(); pixels.width = pixels.height = 0; throw error; }
                preparedEcho = echo;
                return { canvas: pixels, release() {
                  preparingEcho = false;
                  if (echo !== echoPainter) echo?.dispose();
                  pixels.width = pixels.height = 0;
                } };
              },
              onCommit(frame) {
                if (!current() || fit !== token || frame !== queuedFrame || getSnapshot().confirmed !== frame) return;
                const view = getSnapshot();
                echoPainter?.dispose();
                const previous = displayedPainter;
                displayedPainter = framePainter;
                if (previous !== displayedPainter) previous?.dispose();
                paintedFrame = frame.frame;
                paintedView = view;
                firstPainted = true;
                visual = { view, fit: token };
                echoPainter = preparedEcho ?? peerlessPainterFactories.echo(surface, { ...token,
                  current: () => current() && firstPainted && fit === token });
                if (preparedEcho) { preparedEcho.attach(surface); preparedEcho = undefined; }
                else echoPainter.paintEcho(view);
                flushSync(() => {
                  if (!published) { published = true; setPresentation(run!); }
                  notifyVisual();
                });
                signal(); schedule(client);
              },
            });
            const view = getSnapshot();
            try { if (view.confirmed) paintFrame(view.confirmed, view); } catch (error) { report(error); }
          },
        };
        setReadOnly(mode.current);
        setPending(run);
        if (getSnapshot().confirmed === null) { published = true; setPresentation(run); signal(); }
      }).catch(report);
    } catch (error) { report(error); }
    return () => {
      alive = false;
      if (requestMode.current === setReadOnly) requestMode.current = null;
      signal();
      cancelIdle?.();
      for (const dispose of disposers) dispose();
      base?.dispose();
      echoPainter?.dispose();
      framePainter?.dispose();
      if (displayedPainter !== framePainter) displayedPainter?.dispose();
      const ctx = surface.getContext('2d');
      ctx?.setTransform(1, 0, 0, 1, 0, 0);
      ctx?.clearRect(0, 0, surface.width, surface.height);
      activeImages?.release();
      images.dispose();
      pendingMedia.clear();
      for (const face of installedFonts) document.fonts.delete(face);
      clearTimeout(checkpointTimer);
      visualListeners.clear(); failureListeners.clear();
      if (!fatal) void session?.dispose().catch(() => {});
    };
  }, [fileGeneration, canvas]);

  useLayoutEffect(() => {
    const run = presentation?.fileGeneration === fileGeneration ? presentation : null;
    if (!run?.current) return;
    run.setChangeObserver(Boolean(props.onChange));
  }, [presentation, fileGeneration, Boolean(props.onChange)]);

  return { presentation: presentation?.fileGeneration === fileGeneration ? presentation : null,
    pending: pending?.fileGeneration === fileGeneration ? pending : null, failure };
}
