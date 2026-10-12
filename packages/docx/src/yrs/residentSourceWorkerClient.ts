import { ResidentEngineWorkerClient, ResidentWorkerFailureError, takePreloadedResidentEngineWorker,
  type ResidentEngineWorkerPort, type ResidentEngineWorkerFrame } from './residentEngineWorkerClient';
import { isCommandAck } from '../../../../shared/office-session/commands/types';
import { validateResidentSource, type ResidentSourceOpenRequest } from './residentSource';
import type { ResidentFontEnvironmentEvent } from './residentFontEnvironment';
import type { WasmModuleMemory } from '../wasm/loadWasmAsset';
import type { ResidentEngineWorkerResponse } from './residentEngineWorkerProtocol';
import type { Document } from '../types/document';

export interface ResidentEngineWorkerSourceOpened {
  document: Document;
  stateVector: Uint8Array;
  hostJson?: string;
  preview?: ResidentEngineWorkerFrame;
}

export class ResidentSourceWorkerClient extends ResidentEngineWorkerClient {
  private fontListeners = new Set<(event: ResidentFontEnvironmentEvent) => void>();
  private memoryListeners = new Set<(samples: WasmModuleMemory[]) => void>();
  private presentedFontsRevision = 0;

  constructor(worker?: ResidentEngineWorkerPort) {
    super(worker);
    this.observe();
  }

  static adopt(client: ResidentEngineWorkerClient): ResidentSourceWorkerClient {
    if (client instanceof ResidentSourceWorkerClient) return client;
    Object.setPrototypeOf(client, ResidentSourceWorkerClient.prototype);
    const source = client as ResidentSourceWorkerClient;
    source.fontListeners = new Set();
    source.memoryListeners = new Set();
    source.presentedFontsRevision = 0;
    source.observe();
    return source;
  }

  private observe(): void {
    const receive = this.worker.onmessage!;
    this.worker.onmessage = event => {
      receive(event);
      if (this.terminalError) return;
      const response = event.data;
      if (response.ok && response.commandFrame) {
        const pushed = response.commandFrame;
        const identity = this.commandIdentity;
        if (!identity || pushed.session !== identity.session || pushed.incarnation !== identity.incarnation || pushed.epoch !== identity.epoch) return;
      }
      if (response.memory) {
        for (const listener of this.memoryListeners) { try { listener(structuredClone(response.memory)); } catch {} }
      }
      if (response.ok && response.fontEnvironmentEvent) {
        const event = response.fontEnvironmentEvent;
        const identity = this.commandIdentity;
        if (identity && event.session === identity.session && event.incarnation === identity.incarnation && event.epoch === identity.epoch) {
          for (const listener of this.fontListeners) { try { listener(structuredClone(event)); } catch {} }
        }
      }
      if (response.ok) {
        const ack = response.commandAck;
        const identity = this.commandIdentity;
        const valid = identity && ack && isCommandAck(ack) && ack.session === identity.session &&
          ack.incarnation === identity.incarnation && ack.epoch === identity.epoch;
        const frame = valid ? ack.frame : response.commandFrame?.frame;
        const fonts = frame?.payload.fontEnvironment;
        if (frame && fonts && fonts.revision > this.presentedFontsRevision) {
          this.presentedFontsRevision = fonts.revision;
          for (const listener of this.fontListeners) {
            try { listener(structuredClone({ ...fonts, phase: 'presentation', frameEpoch: frame.revision })); } catch {}
          }
        }
      }
    };
  }

  async openSource(source: ResidentSourceOpenRequest): Promise<ResidentEngineWorkerSourceOpened> {
    validateResidentSource(source);
    if (this.openedHeapLimit || this.bootstrapped) {
      throw new ResidentWorkerFailureError('Resident engine worker already holds a document');
    }
    const reservation = { bytes: source.heapLimitBytes };
    this.openedHeapLimit = reservation;
    try {
      const owned: ResidentSourceOpenRequest = source.bytes
        ? { ...source, bytes: source.bytes.slice() } : source;
      const response = await this.request({ type: 'openSource', source: owned }, owned.bytes ? [owned.bytes.buffer] : []);
      if (!response.sourceDocument || !response.stateVector) {
        throw new ResidentWorkerFailureError('Resident engine worker omitted the opened source');
      }
      if (response.layoutRevision !== undefined) this.revision = response.layoutRevision;
      return {
        document: response.sourceDocument, hostJson: response.hostJson,
        stateVector: new Uint8Array(response.stateVector),
        ...(response.frame ? { preview: sourcePreviewFrame(response) } : {}),
      };
    } catch (error) {
      if (this.openedHeapLimit === reservation) this.openedHeapLimit = null;
      throw error;
    }
  }

  async prewarm(): Promise<void> {
    await this.warm();
  }

  onFontEnvironment(listener: (event: ResidentFontEnvironmentEvent) => void): () => void {
    this.fontListeners.add(listener);
    return () => { this.fontListeners.delete(listener); };
  }

  onMemory(listener: (samples: WasmModuleMemory[]) => void): () => void {
    this.memoryListeners.add(listener);
    return () => { this.memoryListeners.delete(listener); };
  }

  override frameRequestPending(reads = true): boolean {
    return [...this.pending.values()].some(({ type }) => type === 'openSource') || super.frameRequestPending(reads);
  }

  protected override request(...args: Parameters<ResidentEngineWorkerClient['request']>) {
    if (args[0].type === 'openSource') this.documentSequence += 1;
    return super.request(...args);
  }

  protected override fenceCommands(): void {
    this.presentedFontsRevision = 0;
    super.fenceCommands();
  }

  protected override fail(error: Error, notify = true): void {
    this.fontListeners?.clear();
    this.memoryListeners?.clear();
    super.fail(error, notify);
  }
}

export function takePreloadedResidentSourceWorker(): ResidentSourceWorkerClient | null {
  const client = takePreloadedResidentEngineWorker();
  return client ? ResidentSourceWorkerClient.adopt(client) : null;
}

function sourcePreviewFrame(
  response: ResidentEngineWorkerResponse & { ok: true }
): ResidentEngineWorkerFrame {
  if (!response.frame)
    throw new ResidentWorkerFailureError('Resident engine worker response omitted its FrameDelta');
  if (!response.caret)
    throw new ResidentWorkerFailureError('Resident engine worker response omitted its caret snapshot');
  if (response.selection === undefined) {
    throw new ResidentWorkerFailureError('Resident engine worker response omitted its selection');
  }
  return {
    frame: new Uint8Array(response.frame),
    ...(response.pageFrames ? { pageFrames: response.pageFrames.map((frame) => new Uint8Array(frame)) } : {}),
    updates: (response.updates ?? []).map((update) => new Uint8Array(update)),
    engineMs: response.engineMs ?? 0,
    workerTotalMs: response.workerTotalMs ?? 0,
    engineProfile: response.engineProfile,
    caret: response.caret,
    selection: response.selection,
    selectionCursor: response.selectionCursor ?? null,
    caretPainted: response.caretPainted ?? false,
    replayMs: response.replayMs ?? 0,
    replayedPages: response.replayedPages ?? 0,
    layoutRevision: response.layoutRevision ?? 0,
    ...(response.documentVersion === undefined ? {} : { documentVersion: response.documentVersion }),
    ...(response.documentPreview === undefined ? {} : { documentPreview: response.documentPreview }),
    ...(response.documentAsOpened === undefined ? {} : { documentAsOpened: response.documentAsOpened }),
    deletedUnits: response.deletedUnits ?? 0,
    ...(response.layoutJson !== undefined ? { layoutJson: response.layoutJson } : {}),
    ...(response.layoutMeta !== undefined ? { layoutMeta: response.layoutMeta } : {}),
    ...(response.layoutProvisional ? { layoutProvisional: true } : {}),
  };
}
