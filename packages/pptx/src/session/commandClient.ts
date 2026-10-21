import { isClientMessage, isHostMessage } from '../../../../shared/office-session/protocol';
import { pipeOpBuffers, type PipeOutputBatch } from '../../../../shared/collaboration-pipe/protocol';
import type { PptxPresencePeer } from '../collaboration/types';
import { createPresentationCollaborationPipe, presentationCollaborationPipes } from './commandCollaboration';
import { createSessionClient } from '../../../../shared/office-session/client';
import type { SessionTransport } from '../../../../shared/office-session/transport';
import { SessionFailure } from '../../../../shared/office-session/types';
import { createWorkerEditSession } from '../../../../shared/office-session/commands/session';
import type { CommandTicket } from '../../../../shared/office-session/commands/types';
import type { PptxFontFace } from '../types';
import { createPresentationSharedAdapter,
  type PresentationSharedAck, type PresentationSharedMethods, type PresentationSharedOpenResult } from './commandSharedAdapter';
import { presentationSharedInputs, type PresentationSharedInput } from './commandSharedInput';
import {
  CommandOutcomeUnknown, PRESENTATION_COMMAND_METHODS,
  type PresentationCommandAck,
  type PresentationCommandBasis,
  type PresentationCommandClient,
  type PresentationCommandEnvelope,
  type PresentationCommandEvents,
  type PresentationCommandFailure,
  type PresentationCommandFrame,
  type PresentationCommandIntent,
  type PresentationCommandMethods,
  type PresentationCommandOpenOptions,
  type PresentationCommandOpenResult,
  type PresentationCommandOutcome,
  type PresentationCommandRefusalReason,
} from './commandProtocol';

type CommandFont = Omit<PptxFontFace, 'bytes'> & { bytes: Uint8Array | ArrayBuffer };
type CommandOpenInput = Omit<PresentationCommandOpenOptions, 'fonts' | 'fallbackFonts'> & {
  fonts?: readonly CommandFont[];
  fallbackFonts?: readonly CommandFont[];
};

export async function createPresentationCommandClient(
  bytes: Uint8Array | ArrayBuffer,
  options: PresentationCommandOpenOptions,
  transport: SessionTransport
): Promise<PresentationCommandClient> {
  return createCommandClient(bytes, options, transport, false);
}

/** @internal */
export async function createSharedPresentationCommandClient(
  bytes: Uint8Array | ArrayBuffer,
  options: CommandOpenInput,
  transport: SessionTransport
): Promise<PresentationCommandClient> {
  return createCommandClient(bytes, options, transport, true);
}

async function createCommandClient(
  bytes: Uint8Array | ArrayBuffer,
  options: CommandOpenInput,
  transport: SessionTransport,
  shared: boolean
): Promise<PresentationCommandClient> {
  const pipe = createPresentationCollaborationPipe();
  type PendingPipeOutput = { batches: PipeOutputBatch<PptxPresencePeer>[]; ready: boolean };
  const pendingPipeOutputs: PendingPipeOutput[] = [];
  const pipeAcks = new WeakMap<object, PendingPipeOutput>();
  let drainingPipe = false;

  function drainPipe(): void {
    if (drainingPipe) return;
    drainingPipe = true;
    try {
      while (pendingPipeOutputs[0]?.ready) pipe.receive(pendingPipeOutputs.shift()!.batches);
    } finally { drainingPipe = false; }
  }

  function receivePipe(batches: PipeOutputBatch<PptxPresencePeer>[], ack?: unknown): void {
    const waitsForAck = ack !== null && typeof ack === 'object' && 'outcomes' in ack;
    const output = { batches, ready: !waitsForAck };
    pendingPipeOutputs.push(output);
    if (waitsForAck) pipeAcks.set(ack, output);
    drainPipe();
  }

  function releasePipe(ack: object): void {
    const output = pipeAcks.get(ack);
    if (output) {
      output.ready = true;
      pipeAcks.delete(ack);
    }
    drainPipe();
  }

  const transportErrors = new Set<(error: unknown) => void>();
  const outstandingCalls = new Set<number>();
  const guarded: SessionTransport = {
    post(message, transfer) {
      const call = isClientMessage(message) && message.kind === 'call' ? message.id : undefined;
      if (call !== undefined) outstandingCalls.add(call);
      try { transport.post(message, transfer); } catch (error) {
        if (call !== undefined) outstandingCalls.delete(call);
        const failure = error instanceof SessionFailure ? error
          : new SessionFailure('crash', error instanceof Error ? error.message : String(error));
        for (const listener of [...transportErrors]) listener(failure);
        throw failure;
      }
    },
    listen(listener) { return transport.listen((message) => {
      if (isHostMessage(message) && (message.kind === 'reply' || message.kind === 'event' && message.name === 'ack')) {
        const output = (message as typeof message & { collabPipe?: PipeOutputBatch<PptxPresencePeer>[] }).collabPipe;
        const awaited = message.kind === 'event' || outstandingCalls.delete(message.id);
        if (output) receivePipe(output, !awaited ? undefined
          : message.kind === 'event' ? message.payload : message.ok ? message.value : undefined);
      }
      listener(message);
    }); },
    onError(listener) {
      transportErrors.add(listener);
      const unsubscribe = transport.onError(listener);
      return () => { transportErrors.delete(listener); unsubscribe(); };
    },
    close() { transport.close(); },
  };
  const rpc = createSessionClient<PresentationCommandMethods & PresentationSharedMethods, PresentationCommandEvents>(guarded, {
    methods: { ...PRESENTATION_COMMAND_METHODS, edit: true, checkpoint: true },
  });
  let state: PresentationCommandOpenResult;
  let adapter: ReturnType<typeof createPresentationSharedAdapter> | undefined;
  let sharedInitial: PresentationSharedAck | undefined;
  try {
    const document = copyBytes(bytes);
    const transfer: Transferable[] = [document];
    const input: PresentationCommandOpenOptions = { ...options,
      fonts: options.fonts?.map((face) => {
        const bytes = copyBytes(face.bytes);
        transfer.push(bytes);
        return { ...face, bytes };
      }),
      fallbackFonts: options.fallbackFonts?.map((face) => {
        const bytes = copyBytes(face.bytes);
        transfer.push(bytes);
        return { ...face, bytes };
      }) };
    if (options.initialUpdate !== undefined) {
      const update = copyBytes(options.initialUpdate);
      input.initialUpdate = new Uint8Array(update);
      transfer.push(update);
    }
    if (options.wasm instanceof ArrayBuffer) {
      input.wasm = options.wasm.slice(0);
      transfer.push(input.wasm);
    }
    state = await rpc.callWithTransfer('openCommandEditor', [document, input], transfer);
    if (rpc.failure) throw rpc.failure;
    const { shared: initial, ...opened } = state as PresentationSharedOpenResult;
    if (shared) {
      adapter = createPresentationSharedAdapter(initial);
      sharedInitial = initial;
    }
    releasePipe(state);
    state = opened;
  } catch (error) {
    await rpc.dispose();
    throw error;
  }
  let sentThroughSeq = state.throughSeq;
  let disposed = false;
  let failure: PresentationCommandFailure | undefined;
  let disposal: Promise<void> | undefined;
  const listeners = new Set<(ack: PresentationCommandAck) => void>();
  const frameListeners = new Set<(frame: PresentationCommandFrame) => void>();
  const mediaListeners = new Set<(media: PresentationCommandFrame['media']) => void>();
  const failureListeners = new Set<(error: PresentationCommandFailure) => void>();
  const updateListeners = new Set<(update: Uint8Array, origin: 'local' | 'remote') => void>();
  let updateThroughSeq = state.throughSeq;
  const sentListeners = new Set<() => void>();
  const media = new Map<string, ArrayBuffer>();
  for (const asset of state.frame?.media ?? []) media.set(asset.assetId, asset.bytes);

  function publishMedia(frame: PresentationCommandFrame | null | undefined): void {
    if (frame?.incarnation !== state.incarnation || frame.documentEpoch !== state.epoch) return;
    const additions = frame.media.filter((asset) => media.get(asset.assetId) !== asset.bytes);
    for (const asset of additions) media.set(asset.assetId, asset.bytes);
    if (additions.length) notify(mediaListeners, additions);
  }

  function notify<T>(callbacks: Set<(value: T) => void>, value: T, ordinary = true): void {
    for (const callback of [...callbacks]) {
      if (disposed || (ordinary && failure)) break;
      if (!callbacks.has(callback)) continue;
      try { callback(value); } catch {}
    }
  }

  function install(ack: PresentationCommandAck): PresentationCommandAck {
    const installed = installAck(ack);
    releasePipe(ack);
    return installed;
  }

  function installAck(ack: PresentationCommandAck): PresentationCommandAck {
    if (disposed || failure || ack.session !== state.session || ack.incarnation !== state.incarnation ||
      ack.epoch !== state.epoch) return ack;
    publishMedia(ack.frame);
    if (disposed || failure) return ack;
    if (ack.throughSeq < state.throughSeq) return ack;
    const advances = ack.throughSeq > state.throughSeq;
    const frame = ack.frame;
    const installsFrame = frame != null && frame.incarnation === state.incarnation &&
      frame.documentEpoch === state.epoch && frame.coversThroughSeq === ack.throughSeq &&
      (frame.slideId === state.activeSlideId || frame.slideId === ack.activeSlideId) &&
      frame.version === ack.version && frame.epoch > (state.frame?.epoch ?? 0);
    state = { ...state, ...ack, frame: installsFrame ? frame : ack.frame === null ? null : state.frame,
      activeSlideId: installsFrame ? frame.slideId : ack.activeSlideId === null ? null : state.activeSlideId };
    if (installsFrame) notify(frameListeners, frame);
    if (advances && !failure) notify(listeners, ack);
    if (ack.throughSeq > updateThroughSeq) {
      updateThroughSeq = ack.throughSeq;
      for (const update of ack.updates ?? []) {
        for (const listener of [...updateListeners]) {
          if (disposed || failure) break;
          if (updateListeners.has(listener)) { try { listener(new Uint8Array(update.bytes).slice(), update.origin); } catch {} }
        }
      }
    }
    return ack;
  }

  function lock(): void {
    state = { ...state, editState: { ...state.editState, readOnly: true,
      toolbar: { ...state.editState.toolbar, canEditText: false } } };
  }

  function fail(error: PresentationCommandFailure): void {
    if (disposed || failure) return;
    failure = error;
    pendingPipeOutputs.length = 0;
    pipe.lose('Collaboration worker failed');
    lock();
    notify(failureListeners, error, false);
  }

  function unavailable(): Error | undefined {
    return disposed ? new Error('Presentation command client is disposed') : failure;
  }

  function reject(error: unknown, seq: number): never {
    if (error instanceof Error && error.name === 'CommandOutcomeUnknown') {
      const refusal: unknown = 'refusal' in error ? error.refusal : undefined;
      const outcomeSeq = refusal !== null && typeof refusal === 'object' && 'seq' in refusal &&
        typeof refusal.seq === 'number' ? refusal.seq : seq;
      fail(new CommandOutcomeUnknown(outcomeSeq, error.message));
    }
    if (rpc.failure) fail(rpc.failure);
    throw failure ?? error;
  }

  const identity = { session: state.session, incarnation: state.incarnation };
  const unack = rpc.on('ack', install);
  const unpipe = rpc.on('collaborationPipe', receivePipe);
  const unfailure = rpc.onFailure(fail);
  const fence = (seq = sentThroughSeq) => ({ ...identity, seq });
  type SharedPayload = { intent: PresentationCommandIntent; basis: PresentationCommandBasis };
  const replies = new Map<number, PresentationCommandAck>();
  const tickets = new Map<number, CommandTicket>();
  let latestSharedAck = sharedInitial;
  const editSession = sharedInitial && adapter ? createWorkerEditSession<
    undefined, SharedPayload, PresentationSharedAck['selection'], PresentationSharedAck['editState'],
    PresentationCommandFrame, never
  >({
    source: undefined,
    connection: {
      initial: sharedInitial,
      client: {
        call: {
          edit(envelope) {
            const error = unavailable();
            if (error) return Promise.reject(error);
            sentThroughSeq = envelope.seq;
            notify(sentListeners, undefined);
            const { basis, intent: draft } = envelope.payload;
            const reference = 'at' in draft.payload ? draft.payload.at : undefined;
            const afterSeq = envelope.afterSeq ?? (reference?.source === 'afterSeq' ? reference.seq : undefined);
            const intent = afterSeq === undefined ? draft : withPredecessor(draft, afterSeq);
            const command = adapter!.envelope({ ...identity, seq: envelope.seq, basis, ...intent,
              ...(afterSeq === undefined ? {} : { afterSeq }) } as PresentationCommandEnvelope);
            const bytes = draft.kind === 'insertPicture' ? draft.payload.bytes.slice(0) : null;
            const reply = bytes ? rpc.callWithTransfer('edit', [{ ...command, payload: { ...command.payload, bytes } }], [bytes])
              : intent.kind === 'collabPipe' ? rpc.callWithTransfer('edit', [command], pipeOpBuffers(intent.payload.ops))
              : rpc.call.edit(command);
            return reply.then((ack) => {
              if (ack.throughSeq >= latestSharedAck!.throughSeq) latestSharedAck = ack;
              const projected = adapter!.ack(ack);
              const output = pipeAcks.get(ack);
              if (output) {
                pipeAcks.set(projected, output);
                pipeAcks.delete(ack);
              }
              replies.set(envelope.seq, projected);
              return ack;
            }, (error: unknown) => reject(error, envelope.seq));
          },
          checkpoint: () => Promise.reject(new Error('Presentation recovery is disabled')),
        },
        get failure() { return rpc.failure; },
        onFailure: (listener) => rpc.onFailure(listener),
        dispose: () => rpc.dispose(),
      },
    },
    recovery: { reopen: () => Promise.reject(new Error('Presentation recovery is disabled')) },
    maxRecoveryAttempts: 0,
    maxInFlight: 256,
    sameCoverageFrames: true,
    discardFrame: (frame) => discardFrameBitmaps(frame.payload),
    echoOverflow: 'decline',
    enqueue: (callback) => callback(),
    requestFullFrame: async () => ({ ...latestSharedAck!, outcomes: [], frame: undefined }),
  }) : undefined;
  const sharedInput: PresentationSharedInput | undefined = editSession && {
    subscribeSent(listener) {
      if (!disposed) sentListeners.add(listener);
      return () => { sentListeners.delete(listener); };
    },
    submit(intent, basis, after) {
      const reference = 'at' in intent.payload ? intent.payload.at : undefined;
      const predecessor = after ?? (reference?.source === 'afterSeq' ? tickets.get(reference.seq) : undefined);
      const ticket = editSession.submit(intent.kind, { intent, basis }, {
        ...(predecessor ? { after: predecessor } : {}),
      });
      if (ticket.seq !== undefined) {
        tickets.set(ticket.seq, ticket);
        while (tickets.size > 256) tickets.delete(tickets.keys().next().value!);
      }
      const operation = ticket.ack.then((outcome) => {
        const ack = ticket.seq === undefined ? undefined : replies.get(ticket.seq);
        if (ticket.seq !== undefined) replies.delete(ticket.seq);
        if (ack) return install(ack);
        const result: PresentationCommandOutcome = { seq: outcome.seq ?? sentThroughSeq + 1,
          status: outcome.status === 'duplicate' ? 'noop' : outcome.status,
          reason: presentationRefusalReason(outcome.reason), committed: false };
        return install({ ...state, frame: undefined, outcomes: [result] });
      }, (error: unknown) => { throw failure ?? error; });
      return { ticket, operation };
    },
  };
  const client: PresentationCommandClient = {
    get state() { return state; },
    dispatch(intent, basis) {
      const error = unavailable();
      if (error) return Promise.reject(error);
      if (sharedInput) return sharedInput.submit(intent,
        basis ?? { epoch: state.epoch, version: state.version }).operation;
      const seq = ++sentThroughSeq;
      const reference = 'at' in intent.payload ? intent.payload.at : undefined;
      const predecessor = reference?.source === 'afterSeq' ? reference.seq : undefined;
      const envelope: PresentationCommandEnvelope = { ...identity, seq,
        basis: basis ?? { epoch: state.epoch, version: state.version },
        ...(predecessor === undefined ? {} : { afterSeq: predecessor }), ...intent };
      const bytes = intent.kind === 'insertPicture' ? intent.payload.bytes.slice(0) : null;
      const reply = bytes ? rpc.callWithTransfer('command', [{ ...envelope, payload: { ...envelope.payload, bytes } } as PresentationCommandEnvelope], [bytes])
        : intent.kind === 'collabPipe' ? rpc.callWithTransfer('command', [envelope], pipeOpBuffers(intent.payload.ops))
        : rpc.call.command(envelope);
      return reply.then(install, (error: unknown) => reject(error, seq));
    },
    async flush(seq) {
      const error = unavailable();
      if (error) return Promise.reject(error);
      if (editSession) await editSession.flush();
      const position = fence(seq);
      return rpc.call.flushCommandEditor(position).then(install, (error: unknown) => reject(error, position.seq));
    },
    async save() {
      const error = unavailable();
      if (error) throw error;
      if (editSession) await editSession.flush();
      const position = fence();
      try {
        const result = await rpc.call.saveCommandEditor(position);
        return new Uint8Array(result.bytes);
      } catch (error) { return reject(error, position.seq); }
    },
    async access(method, ...args) {
      const error = unavailable();
      if (error) throw error;
      if (editSession) await editSession.flush();
      const result = await rpc.call.accessCommandEditor(fence(), method, args);
      const retired = unavailable();
      if (retired) throw retired;
      return result as never;
    },
    async checkpoint() {
      const error = unavailable();
      if (error) throw error;
      if (editSession) await editSession.flush();
      return rpc.call.checkpointCommandEditor(fence());
    },
    async preparePicture(bytes) {
      const error = unavailable();
      if (error) throw error;
      if (editSession) await editSession.flush();
      const owned = bytes.slice(0);
      return rpc.callWithTransfer('prepareCommandPicture', [fence(), owned], [owned]);
    },
    async read(request) {
      const error = unavailable();
      if (error) throw error;
      if (editSession) await editSession.flush();
      const position = fence();
      try {
        const result = await rpc.call.readEditor(position, request);
        install(result);
        return result;
      } catch (error) { return reject(error, position.seq); }
    },
    thumbnail(slideId, knownRevision) {
      const error = unavailable();
      if (error) return Promise.reject(error);
      return rpc.call.thumbnailFrame({ ...identity, slideId, knownRevision })
        .then((reply) => {
          if (disposed || failure || reply.session !== state.session || reply.incarnation !== state.incarnation ||
            reply.epoch !== state.epoch || reply.status !== 'frame') return reply;
          publishMedia(reply.frame);
          return reply;
        })
        .catch((error: unknown) => reject(error, sentThroughSeq));
    },
    subscribe(listener) {
      if (!disposed) listeners.add(listener);
      return () => { listeners.delete(listener); };
    },
    subscribeFrame(listener) {
      if (!disposed) frameListeners.add(listener);
      return () => { frameListeners.delete(listener); };
    },
    subscribeMedia(listener) {
      if (!disposed && !failure) {
        mediaListeners.add(listener);
        if (media.size) notify(new Set([listener]), [...media].map(([assetId, bytes]) => ({ assetId, bytes })));
      }
      return () => { mediaListeners.delete(listener); };
    },
    subscribeUpdate(listener) {
      if (!disposed) updateListeners.add(listener);
      return () => { updateListeners.delete(listener); };
    },
    subscribeFailure(listener) {
      if (!disposed) failureListeners.add(listener);
      return () => { failureListeners.delete(listener); };
    },
    dispose() {
      if (disposal) return disposal;
      disposed = true;
      pendingPipeOutputs.length = 0;
      pipe.lose('Collaboration replica is retired');
      presentationCollaborationPipes.delete(client);
      unpipe();
      unack();
      unfailure();
      listeners.clear();
      frameListeners.clear();
      mediaListeners.clear();
      failureListeners.clear();
      sentListeners.clear();
      updateListeners.clear();
      presentationSharedInputs.delete(client);
      media.clear();
      replies.clear(); tickets.clear();
      disposal = editSession ? editSession.dispose() : rpc.dispose();
      return disposal;
    },
  };
  presentationCollaborationPipes.set(client, pipe);
  if (sharedInput) presentationSharedInputs.set(client, sharedInput);
  return client;
}

function presentationRefusalReason(reason: string | undefined): PresentationCommandRefusalReason | undefined {
  switch (reason) {
    case undefined:
    case 'bad-target':
    case 'invalid-offset':
    case 'stale-basis':
    case 'read-only':
    case 'dependency-refused':
    case 'invalid-composition':
    case 'invalid-payload':
    case 'unsupported':
    case 'unknown-outcome': return reason;
    case 'stale-epoch':
    case 'stale-incarnation':
    case 'wrong-session':
    case 'stale-version': return 'stale-basis';
    default: return 'unknown-outcome';
  }
}

function withPredecessor(intent: PresentationCommandIntent, seq: number): PresentationCommandIntent {
  return { ...intent, payload: { ...intent.payload, at: { source: 'afterSeq', seq } } } as PresentationCommandIntent;
}

function copyBytes(bytes: Uint8Array | ArrayBuffer): ArrayBuffer {
  return ArrayBuffer.isView(bytes)
    ? new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength).slice().buffer
    : bytes.slice(0);
}

function discardFrameBitmaps(payload: unknown): void {
  if (typeof ImageBitmap !== 'function') return;
  const seen = new WeakSet<object>();
  const pending = [payload];
  while (pending.length) {
    const value = pending.pop();
    if (value === null || typeof value !== 'object' || seen.has(value)) continue;
    seen.add(value);
    if (value instanceof ImageBitmap) {
      try { value.close(); } catch {}
    } else if (!(value instanceof ArrayBuffer) && !ArrayBuffer.isView(value)) {
      pending.push(...Object.values(value));
    }
  }
}
