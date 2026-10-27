import { decodeStateVector, parseUpdateMeta } from "yjs";

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

function stripFingerprint(payload: Uint8Array): Uint8Array {
  const suffixLength = FINGERPRINT_TAG.byteLength + 32;
  const tagOffset = payload.byteLength - suffixLength;
  if (tagOffset < 0) return payload;
  for (let index = 0; index < FINGERPRINT_TAG.byteLength; index += 1) {
    if (payload[tagOffset + index] !== FINGERPRINT_TAG[index]) return payload;
  }
  return payload.subarray(0, tagOffset);
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((byte, i) => byte === b[i]);
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
    return this.#applyNow(update);
  }

  #applyNow(update: Uint8Array): ApplyResult {
    const versionBefore = this.#engine.engineVersion();
    let result: unknown;
    try {
      result = this.#engine.applyUpdate(update.slice());
    } catch (error) {
      return rejected(error);
    }
    if (this.#engine.engineVersion() !== versionBefore) {
      this.#version += 1;
      return { accepted: true };
    }
    if (
      this.format === "xlsx" &&
      result !== null &&
      typeof result === "object" &&
      "applied" in result &&
      (result as { applied: unknown }).applied === false &&
      this.#isCausalGap(update)
    ) {
      return { accepted: false, reason: "causalGap" };
    }
    return { accepted: true };
  }

  #isCausalGap(update: Uint8Array): boolean {
    try {
      const meta = parseUpdateMeta(stripFingerprint(update.slice()));
      const serverVector = decodeStateVector(this.#engine.encodeStateVector());
      for (const [client, from] of meta.from) {
        if (from > (serverVector.get(client) ?? 0)) return true;
      }
      return false;
    } catch {
      return false;
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
      let update: Uint8Array | null = null;
      if (!sameBytes(vectorBefore, this.#engine.encodeStateVector())) {
        update = this.#engine.encodeStateAsUpdate(vectorBefore);
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
