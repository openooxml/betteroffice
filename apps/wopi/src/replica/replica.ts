import {
  applyUpdate as applyYjsUpdate,
  Doc as YjsDoc,
  decodeStateVector,
  decodeUpdate,
  encodeStateVector,
  parseUpdateMeta,
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
  applyUpdate(
    update: Uint8Array,
  ): { accepted: true } | { accepted: false; reason: string };
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
  engineVersion(): string;
  dispose(): void;
}

function adaptDocx(session: YrsSession): EngineReplica {
  return {
    clientId: session.clientId,
    encodeStateVector: () => session.encodeStateVector(),
    encodeStateAsUpdate: (remote) => session.encodeStateAsUpdate(remote),
    applyUpdate: (update) => session.applyUpdate(update),
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
    engineVersion: () => handle.version(),
    dispose: () => handle.dispose(),
  };
}

const FINGERPRINT_TAG = new TextEncoder().encode(
  "\0betteroffice-document-fingerprint-v1\0",
);

interface StructLike {
  id: { client: number; clock: number };
  length?: number;
  origin?: { client: number; clock: number } | null;
  rightOrigin?: { client: number; clock: number } | null;
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
  const serverClock = (client: number): number => serverVector.get(client) ?? 0;
  try {
    const meta = parseUpdateMeta(bytes);
    for (const [client, from] of meta.from) {
      if (from > serverClock(client)) return true;
    }
  } catch {}
  let structs: StructLike[] = [];
  let deleteRanges: Array<{ client: number; clock: number; len: number }> = [];
  const readDeleteSet = (deleteSet: unknown): void => {
    const raw = deleteSet as {
      clients?: Map<number, Array<{ clock: number; len: number }>>;
    } | null;
    const clients =
      raw instanceof Map
        ? raw
        : raw?.clients instanceof Map
          ? raw.clients
          : undefined;
    if (!clients) return;
    deleteRanges = [...clients.entries()].flatMap(([client, ranges]) =>
      ranges.map((range) => ({ client, clock: range.clock, len: range.len })),
    );
  };
  try {
    const decoded = decodeUpdate(bytes) as unknown as {
      structs: StructLike[];
      deleteSet: unknown;
    };
    structs = decoded.structs;
    readDeleteSet(decoded.deleteSet);
  } catch {}
  if (structs.length === 0 && deleteRanges.length === 0) {
    try {
      const probe = new YjsDoc();
      applyYjsUpdate(probe, bytes);
      const store = (
        probe as unknown as {
          store: { clients: Map<number, StructLike[]> };
          ds: unknown;
        }
      ).store;
      structs = [...store.clients.values()].flat();
      readDeleteSet((probe as unknown as { ds: unknown }).ds);
    } catch {}
  }
  const provided = new Map<number, Array<[number, number]>>();
  for (const struct of structs) {
    const ranges = provided.get(struct.id.client) ?? [];
    ranges.push([struct.id.clock, struct.id.clock + (struct.length ?? 1)]);
    provided.set(struct.id.client, ranges);
  }
  const covers = (client: number, clock: number): boolean =>
    clock < serverClock(client) ||
    (provided.get(client) ?? []).some(([start, end]) => clock >= start && clock < end);
  const rangeCovered = (client: number, start: number, end: number): boolean => {
    let cursor = start;
    while (cursor < end) {
      if (cursor < serverClock(client)) {
        cursor = Math.max(cursor, Math.min(end, serverClock(client)));
        continue;
      }
      const next = (provided.get(client) ?? [])
        .filter(([from, to]) => from <= cursor && to > cursor)
        .map(([, to]) => to)
        .sort((a, b) => b - a)[0];
      if (next === undefined) return false;
      cursor = next;
    }
    return true;
  };
  for (const struct of structs) {
    for (const ref of [struct.origin, struct.rightOrigin]) {
      if (ref && !covers(ref.client, ref.clock)) return true;
    }
  }
  for (const range of deleteRanges) {
    if (!rangeCovered(range.client, range.clock, range.clock + range.len)) {
      return true;
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

type ApplyResult = { accepted: true } | { accepted: false; reason: string };

class Replica implements ServerReplica {
  readonly format: Format;
  readonly clientId: number;
  #engine: EngineReplica;
  #saveEngine: () => Promise<Uint8Array>;
  #joinPayload: () => JoinPayload;
  #version = 0;
  #destroyed = false;
  #saveChain: Promise<unknown> = Promise.resolve();
  #listeners = new Set<(update: Uint8Array) => void>();

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
    const vectorBefore = this.#engine.encodeStateVector();
    const versionBefore = this.#engine.engineVersion();
    try {
      this.#engine.applyUpdate(update.slice());
    } catch (error) {
      return rejected(error);
    }
    if (this.#engine.engineVersion() !== versionBefore) {
      this.#version += 1;
    }
    this.#emitAuthored(vectorBefore);
    return { accepted: true };
  }

  #authoredUpdate(vectorBefore: Uint8Array): Uint8Array | null {
    const clockBefore = decodeStateVector(vectorBefore.slice()).get(
      this.clientId,
    ) ?? 0;
    const reference = decodeStateVector(
      this.#engine.encodeStateVector().slice(),
    );
    if ((reference.get(this.clientId) ?? 0) === clockBefore) return null;
    reference.set(this.clientId, clockBefore);
    return this.#engine.encodeStateAsUpdate(encodeStateVector(reference));
  }

  #emitAuthored(vectorBefore: Uint8Array): void {
    const update = this.#authoredUpdate(vectorBefore);
    if (!update || this.#destroyed) return;
    for (const listener of [...this.#listeners]) {
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
      const vectorBefore = this.#engine.encodeStateVector();
      const bytes = await this.#saveEngine();
      const update = this.#authoredUpdate(vectorBefore);
      if (update && !this.#destroyed) {
        for (const listener of [...this.#listeners]) {
          try {
            listener(update.slice());
          } catch {}
        }
      }
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
