import {
  decodeStateVector,
  decodeUpdate,
  encodeStateVector,
  mergeUpdates,
  Skip,
  UpdateEncoderV1,
} from "yjs";

import {
  initEngines,
  openDocxReplica,
  openPptxReplica,
  openXlsxReplica,
  saveDocxReplica,
  type PresentationHandle,
  type WorkbookHandle,
  type YrsSession,
} from "./engines";

export type Format = "docx" | "xlsx" | "pptx";

export interface JoinPayload {
  format: Format;
  initialUpdate: Uint8Array;
}

export interface SaveResult {
  bytes: Uint8Array;
  update: Uint8Array | null;
}

export interface ServerReplica {
  readonly format: Format;
  readonly clientId: number;
  readonly version: number;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): ApplyResult;
  joinPayload(): JoinPayload;
  save(): Promise<SaveResult>;
  onLocalUpdate(listener: (update: Uint8Array) => void): () => void;
  destroy(): void;
}

interface EngineReplica {
  readonly clientId: number;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): unknown;
  onLocalUpdate?(listener: (update: Uint8Array) => void): () => void;
  engineVersion(): string;
  dispose(): void;
}

function adaptDocx(session: YrsSession): EngineReplica {
  return {
    clientId: session.clientId,
    encodeStateVector: () => session.encodeStateVector(),
    encodeStateAsUpdate: (remote) => session.encodeStateAsUpdate(remote),
    applyUpdate: (update) => session.applyUpdate(update),
    onLocalUpdate: (listener) => session.onUpdate((update, origin) => {
      if (origin === "local") listener(update);
    }),
    engineVersion: () => session.version(),
    dispose: () => session.destroy(),
  };
}

function adaptWorkbook(handle: WorkbookHandle): EngineReplica {
  return {
    clientId: handle.clientId,
    encodeStateVector: () => handle.encodeStateVector(),
    encodeStateAsUpdate: (remote) => handle.encodeStateAsUpdate(remote),
    applyUpdate: (update) => handle.applyUpdate(update),
    onLocalUpdate: (listener) => handle.onUpdate((update, origin) => {
      if (origin === "local") listener(update);
    }),
    engineVersion: () => handle.version(),
    dispose: () => handle.dispose(),
  };
}

function adaptPresentation(handle: PresentationHandle): EngineReplica {
  return {
    clientId: handle.clientId,
    encodeStateVector: () => handle.encodeStateVector(),
    encodeStateAsUpdate: (remote) => handle.encodeStateAsUpdate(remote),
    applyUpdate: (update) => handle.applyUpdate(update),
    onLocalUpdate: (listener) => handle.onUpdate((update, origin) => {
      if (origin === "local") listener(update);
    }),
    engineVersion: () => handle.version(),
    dispose: () => handle.dispose(),
  };
}

const FINGERPRINT_TAG = new TextEncoder().encode(
  "\0betteroffice-document-fingerprint-v1\0",
);

interface ClockId {
  client: number;
  clock: number;
}

function isClockId(value: unknown): value is ClockId {
  return typeof value === "object" && value !== null &&
    "client" in value && typeof value.client === "number" &&
    "clock" in value && typeof value.clock === "number";
}

/** Whether an update depends on state the server does not have. */
export function hasCausalGap(
  update: Uint8Array,
  serverStateVector: Uint8Array,
): boolean {
  const bytes = stripFingerprint(update.slice());
  let serverVector: Map<number, number>;
  try {
    serverVector = decodeStateVector(serverStateVector.slice());
  } catch {
    return false;
  }
  let decoded: ReturnType<typeof decodeUpdate>;
  try {
    decoded = decodeUpdate(bytes);
  } catch {
    return false;
  }
  const structs = decoded.structs.filter((struct) => !(struct instanceof Skip));
  const provided = new Map<number, Array<[number, number]>>();
  for (const struct of structs) {
    const ranges = provided.get(struct.id.client) ?? [];
    ranges.push([struct.id.clock, struct.id.clock + struct.length]);
    provided.set(struct.id.client, ranges);
  }
  const covered = new Map(serverVector);
  for (const [client, ranges] of provided) {
    let clock = covered.get(client) ?? 0;
    for (const [start, end] of ranges.sort((a, b) => a[0] - b[0])) {
      if (start > clock) return true;
      clock = Math.max(clock, end);
    }
    covered.set(client, clock);
  }
  for (const struct of structs) {
    if (!("origin" in struct)) continue;
    for (const ref of [
      struct.origin,
      struct.rightOrigin,
      struct.parent,
      struct.parentSub,
    ]) {
      if (isClockId(ref) && ref.clock >= (covered.get(ref.client) ?? 0)) {
        return true;
      }
    }
  }
  for (const [client, ranges] of decoded.ds.clients) {
    for (const range of ranges) {
      if (range.clock + range.len > (covered.get(client) ?? 0)) return true;
    }
  }
  return false;
}

function stripFingerprint(payload: Uint8Array): Uint8Array {
  const suffixLength = FINGERPRINT_TAG.byteLength + 32;
  const tagOffset = payload.byteLength - suffixLength;
  if (tagOffset < 0) return payload;
  for (let index = 0; index < FINGERPRINT_TAG.byteLength; index += 1) {
    if (payload[tagOffset + index] !== FINGERPRINT_TAG[index]) return payload;
  }
  return payload.subarray(0, tagOffset);
}

function rejected(error: unknown): { accepted: false; reason: string } {
  const name = error instanceof Error ? error.name : "Error";
  return { accepted: false, reason: `engine rejected update (${name})` };
}

type ApplyResult =
  | { accepted: true; followUps?: Uint8Array[] }
  | { accepted: false; reason: string };

type DeleteRanges = Map<number, Array<{ clock: number; len: number }>>;

interface LocalOperation {
  updates: Uint8Array[];
  checkpoint?: { vector: Uint8Array; deletions: DeleteRanges };
  carried: DeleteRanges;
}

function residualDeletions(
  deletions: DeleteRanges,
  carried: DeleteRanges[],
): DeleteRanges {
  const residual = new Map<number, Array<{ clock: number; len: number }>>();
  for (const [client, ranges] of deletions) {
    const excluded = carried.flatMap((set) => set.get(client) ?? [])
      .sort((a, b) => a.clock - b.clock);
    const remaining: Array<{ clock: number; len: number }> = [];
    for (const range of ranges) {
      let clock = range.clock;
      const end = clock + range.len;
      for (const skip of excluded) {
        if (skip.clock >= end) break;
        if (skip.clock > clock) {
          remaining.push({ clock, len: skip.clock - clock });
        }
        clock = Math.max(clock, skip.clock + skip.len);
        if (clock >= end) break;
      }
      if (clock < end) remaining.push({ clock, len: end - clock });
    }
    if (remaining.length > 0) residual.set(client, remaining);
  }
  return residual;
}

function encodeAuthoredUpdate(
  structs: ReturnType<typeof decodeUpdate>["structs"],
  deletions: DeleteRanges,
): Uint8Array | null {
  if (structs.length === 0 && deletions.size === 0) return null;
  const encoder = new UpdateEncoderV1();
  const clients = new Map<number, typeof structs>();
  for (const struct of structs) {
    const items = clients.get(struct.id.client) ?? [];
    items.push(struct);
    clients.set(struct.id.client, items);
  }
  encoder.writeLen(clients.size);
  for (const [client, items] of [...clients].sort((a, b) => b[0] - a[0])) {
    encoder.writeLen(items.length);
    encoder.writeClient(client);
    encoder.writeLen(items[0]!.id.clock);
    for (const item of items) item.write(encoder, 0);
  }
  encoder.writeLen(deletions.size);
  for (const [client, ranges] of [...deletions].sort((a, b) => b[0] - a[0])) {
    encoder.resetDsCurVal();
    encoder.writeClient(client);
    encoder.writeLen(ranges.length);
    for (const range of ranges) {
      encoder.writeDsClock(range.clock);
      encoder.writeDsLen(range.len);
    }
  }
  return encoder.toUint8Array();
}

export class Replica implements ServerReplica {
  readonly format: Format;
  readonly clientId: number;
  #engine: EngineReplica;
  #saveEngine: () => Promise<Uint8Array>;
  #joinPayload: () => JoinPayload;
  #version = 0;
  #destroyed = false;
  #saveChain: Promise<unknown> = Promise.resolve();
  #listeners = new Set<(update: Uint8Array) => void>();
  #operations: LocalOperation[] = [];
  #unsubscribe?: () => void;

  constructor(
    format: Format,
    engine: EngineReplica,
    saveEngine: () => Promise<Uint8Array>,
    joinPayload: () => JoinPayload,
  ) {
    this.format = format;
    this.clientId = engine.clientId;
    this.#engine = engine;
    this.#saveEngine = saveEngine;
    this.#joinPayload = joinPayload;
    this.#unsubscribe = engine.onLocalUpdate?.((update) => {
      const operation = this.#operations.at(-1);
      if (operation) operation.updates.push(update.slice());
      else this.#publish(update);
    });
  }

  get version(): number {
    return this.#version;
  }

  #assertAlive(): void {
    if (this.#destroyed) throw new Error("server replica is destroyed");
  }

  encodeStateVector(): Uint8Array {
    this.#assertAlive();
    return this.#engine.encodeStateVector();
  }

  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array {
    this.#assertAlive();
    return this.#engine.encodeStateAsUpdate(remoteStateVector);
  }

  applyUpdate(update: Uint8Array): ApplyResult {
    this.#assertAlive();
    if (update.length === 0) return { accepted: true };
    return this.#applyNow(stripFingerprint(update));
  }

  #applyNow(update: Uint8Array): ApplyResult {
    if (hasCausalGap(update, this.#engine.encodeStateVector())) {
      return { accepted: false, reason: "causalGap" };
    }
    const versionBefore = this.#engine.engineVersion();
    const operation = this.#beginOperation(update);
    try {
      this.#engine.applyUpdate(update.slice());
    } catch (error) {
      this.#endOperation(operation, false);
      return rejected(error);
    }
    if (this.#engine.engineVersion() !== versionBefore) {
      this.#version += 1;
    }
    const authored = this.#endOperation(operation);
    if (!authored) return { accepted: true };
    this.#publish(authored);
    return { accepted: true, followUps: [authored] };
  }

  #checkpoint(): NonNullable<LocalOperation["checkpoint"]> {
    const vector = this.#engine.encodeStateVector();
    const deletions = decodeUpdate(stripFingerprint(
      this.#engine.encodeStateAsUpdate(vector),
    )).ds.clients;
    return { vector, deletions };
  }

  #captureFallback(operation: LocalOperation): void {
    if (!operation.checkpoint) return;
    const { vector, deletions } = operation.checkpoint;
    const clockBefore = decodeStateVector(vector.slice()).get(
      this.clientId,
    ) ?? 0;
    const reference = decodeStateVector(
      this.#engine.encodeStateVector().slice(),
    );
    reference.set(this.clientId, clockBefore);
    const decoded = decodeUpdate(stripFingerprint(
      this.#engine.encodeStateAsUpdate(encodeStateVector(reference)),
    ));
    const update = encodeAuthoredUpdate(
      decoded.structs,
      residualDeletions(decoded.ds.clients, [deletions, operation.carried]),
    );
    if (update) operation.updates.push(update);
  }

  #beginOperation(carried?: Uint8Array): LocalOperation {
    const operation: LocalOperation = { updates: [], carried: new Map() };
    if (!this.#unsubscribe) {
      const parent = this.#operations.at(-1);
      if (parent) this.#captureFallback(parent);
      operation.checkpoint = this.#checkpoint();
      if (carried) {
        try {
          operation.carried = decodeUpdate(carried).ds.clients;
        } catch {}
      }
    }
    this.#operations.push(operation);
    return operation;
  }

  #endOperation(operation: LocalOperation, accepted = true): Uint8Array | null {
    try {
      if (accepted) this.#captureFallback(operation);
    } finally {
      this.#operations.pop();
      const parent = this.#operations.at(-1);
      if (parent && !this.#unsubscribe) parent.checkpoint = this.#checkpoint();
    }
    if (!accepted || operation.updates.length === 0) return null;
    return operation.updates.length === 1
      ? operation.updates[0]!
      : mergeUpdates(operation.updates);
  }

  #publish(update: Uint8Array): void {
    if (this.#destroyed) return;
    for (const listener of [...this.#listeners]) {
      if (this.#destroyed) return;
      try {
        listener(update.slice());
      } catch {}
    }
  }

  joinPayload(): JoinPayload {
    this.#assertAlive();
    return this.#joinPayload();
  }

  onLocalUpdate(listener: (update: Uint8Array) => void): () => void {
    this.#assertAlive();
    this.#listeners.add(listener);
    return () => {
      this.#listeners.delete(listener);
    };
  }

  save(): Promise<SaveResult> {
    this.#assertAlive();
    const run = this.#saveChain.then(async (): Promise<SaveResult> => {
      this.#assertAlive();
      const operation = this.#beginOperation();
      let bytes: Uint8Array;
      try {
        bytes = await this.#saveEngine();
      } catch (error) {
        const update = this.#endOperation(operation);
        if (update) this.#publish(update);
        throw error;
      }
      const update = this.#endOperation(operation);
      if (update) this.#publish(update);
      return { bytes, update };
    });
    this.#saveChain = run.then(
      () => undefined,
      () => undefined,
    );
    return run;
  }

  destroy(): void {
    this.#assertAlive();
    this.#destroyed = true;
    void this.#saveChain.then(() => {
      try {
        this.#unsubscribe?.();
        this.#engine.dispose();
      } catch {}
    });
  }
}

/** Opens the server replica for one file. The server is the first replica. */
export async function openServerReplica(
  format: Format,
  bytes: Uint8Array,
  options: { clientId: number },
): Promise<ServerReplica> {
  await initEngines();
  if (format === "docx") {
    const session = await openDocxReplica(bytes, options.clientId);
    return new Replica(
      format,
      adaptDocx(session),
      () => saveDocxReplica(session),
      () => ({
        format: "docx",
        initialUpdate: session.encodeStateAsUpdate(),
      }),
    );
  }
  if (format === "xlsx") {
    const handle = openXlsxReplica(bytes, options.clientId);
    return new Replica(
      format,
      adaptWorkbook(handle),
      async () => handle.save(),
      () => ({ format: "xlsx", initialUpdate: handle.encodeStateAsUpdate() }),
    );
  }
  const handle = openPptxReplica(bytes, options.clientId);
  return new Replica(
    format,
    adaptPresentation(handle),
    async () => handle.save(),
    () => ({ format: "pptx", initialUpdate: handle.encodeStateAsUpdate() }),
  );
}
