import { useEffect, useRef, useState } from 'react';
import { ResidentWorkerOutOfMemoryError } from '@betteroffice/docx/yrs';
import { createDocxWorkerEditSession, discardCommandPresentation, SessionFailure } from '@betteroffice/docx/worker';
import type {
  DocxCommandFrame,
  DocxCommandPageIdentity,
  DocxWorkerEditSession,
  DocxWorkerEditSessionOptions,
  DocxCommandConnection,
  DocxCommandViewportRequest,
} from '@betteroffice/docx/worker';
import { createPeerlessCommandBinding, type PeerlessCommandBinding } from '../peerless/commandBinding';
import { createEchoPainter } from '../peerless/echoPainter';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from '../peerless/echoPainter';
import { createFramePresentation, type PeerlessFramePresentation } from '../peerless/framePresentation';
import { coalesceInputDrafts, holdInputForCoalescing } from '../peerless/inputCoalescing';
import type { Document } from '@betteroffice/docx/types/document';
import { createInputAdmission, type PeerlessInputAdmission } from '../peerless/inputAdmission';
import { createPeerlessViewStore, type PeerlessFrameGeometry, type PeerlessProvisionalCaret, type PeerlessViewStore } from '../peerless/viewStore';

type ResidentEdit = DocxCommandConnection['client']['call']['edit'];

function replayEdits(resident: DocxCommandConnection, appliedSeqs: readonly number[], onSkipped: () => void): ResidentEdit {
  const skippedVersions: number[] = [];
  const skippedPredecessors = new Map<number, number | undefined>();
  let previousSeq = resident.initial.throughSeq || undefined;
  return async command => {
    if (command.afterSeq !== undefined && skippedPredecessors.has(command.afterSeq)) {
      command = { ...command, afterSeq: skippedPredecessors.get(command.afterSeq) };
    }
    if (command.replay && skippedVersions.length) {
      const skipped = skippedVersions.filter(version => version <= command.basis.version).length;
      command = { ...command, basis: { ...command.basis, version: command.basis.version - skipped } };
    }
    const ack = await resident.client.call.edit(command);
    if (!command.replay) return ack;
    const skipped = ack.outcomes.find(outcome => command.kind === 'replacePasteFragment' && outcome.seq === command.seq &&
      outcome.status === 'refused' && (outcome.reason === 'paste-changed' || outcome.reason === 'nothing-to-undo'));
    const failed = ack.outcomes.find(outcome => outcome !== skipped && (outcome.status === 'refused' || outcome.status === 'cancelled' ||
      (outcome.status === 'noop' && appliedSeqs.includes(outcome.seq))));
    if (failed) {
      discardCommandPresentation(ack.frame);
      throw new SessionFailure('message', `DOCX recovery failed at command ${failed.seq} (${command.kind}): ${failed.reason ?? failed.status}`);
    }
    if (skipped) {
      // Later edits may depend on the choice's sequence or have observed its version.
      skippedPredecessors.set(command.seq, previousSeq);
      if (appliedSeqs.includes(command.seq)) skippedVersions.push(ack.version + skippedVersions.length + 1);
      onSkipped();
      return { ...ack, outcomes: ack.outcomes.map(outcome => outcome === skipped ? { seq: outcome.seq, status: 'noop' as const } : outcome) };
    }
    previousSeq = command.seq;
    return ack;
  };
}

/** @internal Opens worker authority without creating a main-thread document. */
export type PeerlessDocumentSource = Uint8Array | Document;

export interface UsePeerlessEditSessionOptions {
  source: PeerlessDocumentSource;
  rasterScale?: number;
  openConnection(source: PeerlessDocumentSource, signal: AbortSignal, getRasterScale?: () => number): Promise<DocxCommandConnection>;
  recovery?: DocxWorkerEditSessionOptions<PeerlessDocumentSource, PeerlessEchoIntent>['recovery'];
  canvasForPage(page: DocxCommandPageIdentity): HTMLCanvasElement | null;
  prepareGeometry(frame: DocxCommandFrame, rasterScale?: number): PeerlessFrameGeometry | undefined;
  /** The same frame geometry with page boxes snapped to another raster scale. */
  rescaleGeometry?(geometry: PeerlessFrameGeometry, rasterScale: number): PeerlessFrameGeometry | undefined;
  onError?(error: Error, recoveryFailed?: boolean): void;
  /** Failures the editor survives, such as pages that cannot be painted until the viewport or zoom changes. */
  onRecoverableError?(error: Error): void;
  onPasteOptionNotRestored?(): void;
}

/** @internal */
export interface PeerlessEditSessionBinding {
  session: DocxWorkerEditSession<PeerlessEchoIntent>;
  viewStore: PeerlessViewStore;
  admission: PeerlessInputAdmission;
  commands: PeerlessCommandBinding;
  echoPainter: PeerlessEchoPainter;
  presentation: PeerlessFramePresentation;
  setPasteOptionsLive?(live: boolean): void;
  setDisplayWindow?(start: number, end: number): void;
  setRasterScale?(scale: number): void;
}

/** @internal */
export type UsePeerlessEditSessionResult =
  | { status: 'opening'; binding: null; error: null }
  | { status: 'ready'; binding: PeerlessEditSessionBinding; error: null }
  | { status: 'failed'; binding: null; error: Error; recoveryFailed: boolean };

/** @internal */
export function usePeerlessEditSession(
  options: UsePeerlessEditSessionOptions
): UsePeerlessEditSessionResult {
  const [result, setResult] = useState<UsePeerlessEditSessionResult>({ status: 'opening', binding: null, error: null });
  const callbacks = useRef(options);
  callbacks.current = options;
  const { source, rasterScale = 1 } = options;
  const appliedRasterScale = useRef(rasterScale);
  const pasteOptionsLive = useRef(false);
  useEffect(() => {
    pasteOptionsLive.current = false;
    const controller = new AbortController();
    let binding: PeerlessEditSessionBinding | undefined;
    let connection: DocxCommandConnection | undefined;
    let unsubscribe: (() => void) | undefined;
    let incarnation = 0;
    let sessionCreated = false;
    let respawned = false;
    let recovering = false;
    let pasteOptionSkipped = false;
    let crash: Error | undefined;
    let receiveRequestedFrame: ((frame: DocxCommandFrame) => void) | undefined;
    const connect = (resident: DocxCommandConnection, replayAppliedSeqs?: readonly number[]): DocxCommandConnection => ({
      ...resident,
      client: {
        ...resident.client,
        get failure() { return resident.client.failure; },
        call: replayAppliedSeqs ? { ...resident.client.call,
          edit: replayEdits(resident, replayAppliedSeqs, () => { pasteOptionSkipped = true; }),
        } : resident.client.call,
        onFrame(listener) {
          receiveRequestedFrame = listener;
          const stop = resident.client.onFrame?.(listener);
          return () => {
            if (receiveRequestedFrame === listener) receiveRequestedFrame = undefined;
            stop?.();
          };
        },
      },
    });
    const cleanups: (() => void)[] = [];
    const close = () => {
      unsubscribe?.();
      for (const cleanup of cleanups.splice(0).reverse()) cleanup();
      if (!sessionCreated && connection) {
        discardCommandPresentation(connection.initial.frame);
        void connection.client.dispose().catch(() => {});
      }
    };
    const fail = (cause: unknown, recoveryFailed = false) => {
      if (controller.signal.aborted) return;
      const error = cause instanceof Error ? cause : new Error(String(cause));
      controller.abort();
      close();
      setResult({ status: 'failed', binding: null, error, recoveryFailed });
      callbacks.current.onError?.(error, recoveryFailed);
    };
    setResult({ status: 'opening', binding: null, error: null });
    void (async () => {
      try {
        const getRasterScale = () => {
          appliedRasterScale.current = callbacks.current.rasterScale ?? 1;
          return appliedRasterScale.current;
        };
        getRasterScale();
        connection = await callbacks.current.openConnection(source, controller.signal, getRasterScale);
        if (controller.signal.aborted) {
          close();
          return;
        }
        const initial = connection.initial;
        incarnation = initial.incarnation;
        const viewStore = createPeerlessViewStore({
          identity: { session: initial.session, incarnation, epoch: initial.epoch },
          version: initial.version, throughSeq: initial.throughSeq, presentedThroughSeq: -1,
          workerSelection: initial.selection, anchors: null, editState: initial.editState, geometry: null,
        });
        cleanups.push(() => viewStore.dispose());
        const canvasForPage = (page: DocxCommandPageIdentity) => callbacks.current.canvasForPage(page);
        const requestFullFrame = (settled: () => void = () => {}) => {
          if (!connection || controller.signal.aborted) { settled(); return; }
          const resident = connection;
          void resident.requestFullFrame().then((ack) => {
            settled();
            if (!ack.frame) return;
            if (!controller.signal.aborted && resident === connection && receiveRequestedFrame) receiveRequestedFrame(ack.frame);
            else discardCommandPresentation(ack.frame);
          }).catch((error: unknown) => {
            settled();
            if (!controller.signal.aborted && !recovering) callbacks.current.onError?.(error instanceof Error ? error : new Error(String(error)));
          });
        };
        let echoCaret: PeerlessProvisionalCaret | undefined;
        const clearEchoCaret = () => {
          if (echoCaret && viewStore.getSnapshot().provisionalCaret === echoCaret) viewStore.setProvisionalCaret(null);
          echoCaret = undefined;
        };
        const echoPainter = createEchoPainter({ canvasForPage, requestFullFrame, onPaint: (predicted) => {
          if (!predicted) { clearEchoCaret(); return; }
          const bounds = viewStore.getSnapshot().geometry?.pages.find((item) => item.page.pageIndex === predicted.pageIndex)?.bounds;
          if (!bounds) { clearEchoCaret(); return; }
          const caret: PeerlessProvisionalCaret = { source: 'echo', get throughSeq() { return predicted.throughSeq; },
            pageIndex: predicted.pageIndex,
            rect: { x: predicted.rect.x + bounds.x, y: predicted.rect.y + bounds.y, width: predicted.rect.width, height: predicted.rect.height } };
          viewStore.setProvisionalCaret(caret);
          const shown = viewStore.getSnapshot().provisionalCaret;
          echoCaret = shown === caret ? caret : shown === echoCaret ? echoCaret : undefined;
        } });
        cleanups.push(() => echoPainter.dispose());
        const presentation = createFramePresentation({
          viewStore, echoPainter, canvasForPage, prepareGeometry: (frame) => callbacks.current.prepareGeometry(frame),
          requestFullFrame,
          onError: (error) => { if (!controller.signal.aborted) callbacks.current.onRecoverableError?.(error); },
        });
        cleanups.push(() => presentation.dispose());
        let displayWindow = { start: 0, end: 3 };
        let viewport: Omit<DocxCommandViewportRequest, 'frameEpoch'> | undefined;
        let viewportPending = false;
        const requestViewport = async () => {
          if (viewportPending || !connection?.requestViewport || controller.signal.aborted) return;
          viewportPending = true;
          try {
            while (viewport && !controller.signal.aborted) {
              const request = viewport;
              viewport = undefined;
              await connection.requestViewport?.({ ...request, frameEpoch: viewStore.getSnapshot().geometry?.frameEpoch ?? 0 });
            }
          } catch (error) {
            if (!controller.signal.aborted && !recovering) callbacks.current.onError?.(error instanceof Error ? error : new Error(String(error)));
          } finally { viewportPending = false; }
        };
        let checkpointDeferred = false;
        const session = createDocxWorkerEditSession<PeerlessDocumentSource, PeerlessEchoIntent>({
          source, connection: connect(connection), echo: presentation.hooks, coalesce: coalesceInputDrafts,
          holdForCoalescing: holdInputForCoalescing,
          transactionalPresentation: true, echoOverflow: 'decline', maxEchoEntries: 256,
          maxRecoveryAttempts: callbacks.current.recovery ? 1 : 0,
          checkpointPolicy: ({ applied, bytes }) => {
            const due = applied >= 128 || bytes >= 1024 * 1024;
            checkpointDeferred = due && pasteOptionsLive.current;
            return due && !checkpointDeferred;
          },
          recovery: {
            checkpoint: callbacks.current.recovery?.checkpoint
              ? (resident) => callbacks.current.recovery?.checkpoint?.(resident) ?? resident.client.call.checkpoint() : undefined,
            onRecovery(event) {
              if (controller.signal.aborted) return;
              if (event.phase === 'failed') fail(event.failure ?? new Error('The worker editor session failed'), recovering);
              if (event.phase !== 'diverged') recovering = event.phase === 'recovering';
              if (event.phase === 'recovering') {
                crash = event.failure;
                viewStore.setRecovering?.(true);
              }
              if (event.phase === 'recovered') {
                viewport = { ...displayWindow, ...viewport };
                void requestViewport();
                if (pasteOptionSkipped) callbacks.current.onPasteOptionNotRestored?.();
              }
              callbacks.current.recovery?.onRecovery?.(event);
            },
            async reopen(input) {
              const recovery = callbacks.current.recovery;
              if (!recovery) throw new Error('Worker editor recovery is unavailable');
              if (respawned) throw crash ?? new Error('The worker editor session failed');
              respawned = true;
              const reopened = await recovery.reopen(input).catch((error: unknown) => {
                if (error instanceof ResidentWorkerOutOfMemoryError) {
                  throw new SessionFailure('out-of-memory', error.message, JSON.stringify(error.memory));
                }
                throw error;
              });
              if (controller.signal.aborted) {
                discardCommandPresentation(reopened.initial.frame);
                void reopened.client.dispose().catch(() => {});
                throw new Error('The worker editor session was closed');
              }
              connection = reopened;
              incarnation = reopened.initial.incarnation;
              viewStore.acceptWorkerState({ session: reopened.initial.session, incarnation, epoch: reopened.initial.epoch }, {
                epoch: reopened.initial.epoch, version: reopened.initial.version, throughSeq: reopened.initial.throughSeq,
                presentedThroughSeq: -1, selection: reopened.initial.selection, editState: reopened.initial.editState,
              });
              return connect(reopened, input.replayAppliedSeqs);
            },
          },
        });
        sessionCreated = true;
        cleanups.push(() => { void session.dispose().catch(() => {}); });
        if (controller.signal.aborted) {
          close();
          return;
        }
        const admission = createInputAdmission({ session, viewStore, echoPainter });
        cleanups.push(() => admission.dispose());
        binding = { session, viewStore, admission, commands: createPeerlessCommandBinding(admission), echoPainter, presentation,
          setPasteOptionsLive(live) {
            if (controller.signal.aborted || pasteOptionsLive.current === live) return;
            pasteOptionsLive.current = live;
            if (!live && checkpointDeferred) void session.flush().catch(() => {});
          },
          setDisplayWindow(start, end) {
            displayWindow = { start, end };
            viewport = { ...viewport, ...displayWindow };
            void requestViewport();
            presentation.invalidate();
          },
          setRasterScale(scale) {
            viewport = { ...displayWindow, ...viewport, rasterScale: scale };
            void requestViewport();
            presentation.invalidate();
          },
        };
        unsubscribe = session.subscribe((state) => viewStore.acceptWorkerState({ session: initial.session, incarnation, epoch: state.epoch }, state));
        viewStore.acceptWorkerState({ session: initial.session, incarnation, epoch: session.state.epoch }, session.state);
        setResult({ status: 'ready', binding, error: null });
      } catch (cause) {
        fail(cause);
      }
    })();
    return () => {
      controller.abort();
      close();
    };
  }, [source]);
  useEffect(() => {
    if (result.status !== 'ready' || rasterScale === appliedRasterScale.current) return;
    const timer = setTimeout(() => {
      appliedRasterScale.current = rasterScale;
      result.binding.setRasterScale?.(rasterScale);
    }, 150);
    return () => clearTimeout(timer);
  }, [result.binding, rasterScale]);
  return result;
}
