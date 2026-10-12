import { useEffect, useRef, useState } from 'react';
import { createDocxWorkerEditSession, discardCommandPresentation } from '@betteroffice/docx/peerless';
import type {
  DocxCommandFrame,
  DocxCommandPageIdentity,
  DocxWorkerEditSession,
  DocxWorkerEditSessionOptions,
  ResidentCommandConnection,
} from '@betteroffice/docx/peerless';
import { createPeerlessCommandBinding, type PeerlessCommandBinding } from '../peerless/commandBinding';
import { createEchoPainter } from '../peerless/echoPainter';
import type { PeerlessEchoIntent, PeerlessEchoPainter } from '../peerless/echoPainter';
import { createFramePresentation, type PeerlessFramePresentation } from '../peerless/framePresentation';
import { coalesceInputDrafts } from '../peerless/inputCoalescing';
import type { Document } from '@betteroffice/docx/types/document';
import { createInputAdmission, type PeerlessInputAdmission } from '../peerless/inputAdmission';
import { createPeerlessViewStore, type PeerlessFrameGeometry, type PeerlessViewStore } from '../peerless/viewStore';

/** @internal Opens worker authority without creating a main-thread document. */
export type PeerlessDocumentSource = Uint8Array | Document;

export interface UsePeerlessEditSessionOptions {
  source: PeerlessDocumentSource;
  openConnection(source: PeerlessDocumentSource, signal: AbortSignal): Promise<ResidentCommandConnection>;
  recovery?: DocxWorkerEditSessionOptions<PeerlessDocumentSource, PeerlessEchoIntent>['recovery'];
  canvasForPage(page: DocxCommandPageIdentity): HTMLCanvasElement | null;
  prepareGeometry(frame: DocxCommandFrame): PeerlessFrameGeometry | undefined;
  canEchoFont?(family: string): boolean;
  onError?(error: Error): void;
}

/** @internal */
export interface PeerlessEditSessionBinding {
  session: DocxWorkerEditSession<PeerlessEchoIntent>;
  viewStore: PeerlessViewStore;
  admission: PeerlessInputAdmission;
  commands: PeerlessCommandBinding;
  echoPainter: PeerlessEchoPainter;
  presentation: PeerlessFramePresentation;
  setDisplayWindow?(start: number, end: number): void;
}

/** @internal */
export type UsePeerlessEditSessionResult =
  | { status: 'opening'; binding: null; error: null }
  | { status: 'ready'; binding: PeerlessEditSessionBinding; error: null }
  | { status: 'failed'; binding: null; error: Error };

/** @internal */
export function usePeerlessEditSession(
  options: UsePeerlessEditSessionOptions
): UsePeerlessEditSessionResult {
  const [result, setResult] = useState<UsePeerlessEditSessionResult>({ status: 'opening', binding: null, error: null });
  const callbacks = useRef(options);
  callbacks.current = options;
  const { source } = options;
  useEffect(() => {
    const controller = new AbortController();
    let binding: PeerlessEditSessionBinding | undefined;
    let connection: ResidentCommandConnection | undefined;
    let unsubscribe: (() => void) | undefined;
    let incarnation = 0;
    let sessionCreated = false;
    let receiveRequestedFrame: ((frame: DocxCommandFrame) => void) | undefined;
    const connect = (resident: ResidentCommandConnection): ResidentCommandConnection => ({
      ...resident,
      client: {
        ...resident.client,
        get failure() { return resident.client.failure; },
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
    const fail = (cause: unknown) => {
      if (controller.signal.aborted) return;
      const error = cause instanceof Error ? cause : new Error(String(cause));
      controller.abort();
      close();
      setResult({ status: 'failed', binding: null, error });
      callbacks.current.onError?.(error);
    };
    setResult({ status: 'opening', binding: null, error: null });
    void (async () => {
      try {
        connection = await callbacks.current.openConnection(source, controller.signal);
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
        const echoPainter = createEchoPainter({ canvasForPage, canEchoFont: family => callbacks.current.canEchoFont?.(family) ?? true });
        cleanups.push(() => echoPainter.dispose());
        const presentation = createFramePresentation({
          viewStore, echoPainter, canvasForPage, prepareGeometry: (frame) => callbacks.current.prepareGeometry(frame),
          requestFullFrame: () => {
            if (!connection || controller.signal.aborted) return;
            const resident = connection;
            void resident.requestFullFrame().then((ack) => {
              if (!ack.frame) return;
              if (!controller.signal.aborted && resident === connection && receiveRequestedFrame) receiveRequestedFrame(ack.frame);
              else discardCommandPresentation(ack.frame);
            }).catch((error: unknown) => {
              if (!controller.signal.aborted) callbacks.current.onError?.(error instanceof Error ? error : new Error(String(error)));
            });
          },
        });
        cleanups.push(() => presentation.dispose());
        const session = createDocxWorkerEditSession<PeerlessDocumentSource, PeerlessEchoIntent>({
          source, connection: connect(connection), echo: presentation.hooks, coalesce: coalesceInputDrafts,
          transactionalPresentation: true, echoOverflow: 'decline', maxEchoEntries: 256,
          maxRecoveryAttempts: callbacks.current.recovery ? undefined : 0,
          recovery: {
            checkpoint: callbacks.current.recovery?.checkpoint
              ? (resident) => callbacks.current.recovery?.checkpoint?.(resident) ?? resident.client.call.checkpoint() : undefined,
            onRecovery(event) {
              if (controller.signal.aborted) return;
              if (event.phase === 'failed') fail(event.failure ?? new Error('Peerless session failed'));
              callbacks.current.recovery?.onRecovery?.(event);
            },
            async reopen(input) {
              const recovery = callbacks.current.recovery;
              if (!recovery) throw new Error('Peerless recovery is unavailable');
              const reopened = await recovery.reopen(input);
              if (controller.signal.aborted) {
                discardCommandPresentation(reopened.initial.frame);
                void reopened.client.dispose().catch(() => {});
                throw new Error('Peerless session was closed');
              }
              connection = reopened;
              incarnation = reopened.initial.incarnation;
              viewStore.acceptWorkerState({ session: reopened.initial.session, incarnation, epoch: reopened.initial.epoch }, {
                epoch: reopened.initial.epoch, version: reopened.initial.version, throughSeq: reopened.initial.throughSeq,
                presentedThroughSeq: -1, selection: reopened.initial.selection, editState: reopened.initial.editState,
              });
              return connect(reopened);
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
        let viewport: { start: number; end: number } | undefined;
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
            if (!controller.signal.aborted) callbacks.current.onError?.(error instanceof Error ? error : new Error(String(error)));
          } finally { viewportPending = false; }
        };
        binding = { session, viewStore, admission, commands: createPeerlessCommandBinding(admission), echoPainter, presentation,
          setDisplayWindow(start, end) {
            viewport = { start, end };
            void requestViewport();
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
  return result;
}
