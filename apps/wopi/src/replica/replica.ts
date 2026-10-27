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
  save(): Promise<Uint8Array>;
  destroy(): void;
}

interface EngineReplica {
  readonly clientId: number;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): unknown;
  dispose(): void;
}

function adaptDocx(session: YrsSession): EngineReplica {
  return {
    clientId: session.clientId,
    encodeStateVector: () => session.encodeStateVector(),
    encodeStateAsUpdate: (remote) => session.encodeStateAsUpdate(remote),
    applyUpdate: (update) => session.applyUpdate(update),
    dispose: () => session.destroy(),
  };
}

function adaptWorkbook(handle: WorkbookHandle): EngineReplica {
  return {
    clientId: handle.clientId,
    encodeStateVector: () => handle.encodeStateVector(),
    encodeStateAsUpdate: (remote) => handle.encodeStateAsUpdate(remote),
    applyUpdate: (update) => handle.applyUpdate(update),
    dispose: () => handle.dispose(),
  };
}

function adaptPresentation(handle: PresentationHandle): EngineReplica {
  return {
    clientId: handle.clientId,
    encodeStateVector: () => handle.encodeStateVector(),
    encodeStateAsUpdate: (remote) => handle.encodeStateAsUpdate(remote),
    applyUpdate: (update) => handle.applyUpdate(update),
    dispose: () => handle.dispose(),
  };
}

function isEmptyUpdate(update: Uint8Array): boolean {
  return update.length === 0 || (update.length === 2 && update[0] === 0 && update[1] === 0);
}

function rejected(error: unknown): { accepted: false; reason: string } {
  const name = error instanceof Error ? error.name : "Error";
  return { accepted: false, reason: `engine rejected update (${name})` };
}

class Replica implements ServerReplica {
  readonly format: Format;
  readonly clientId: number;
  #engine: EngineReplica;
  #saveEngine: () => Promise<Uint8Array>;
  #joinPayload: () => JoinPayload;
  #version = 0;
  #destroyed = false;
  #saveChain: Promise<unknown> = Promise.resolve();
  #notAppliedIsNoOp: boolean;

  constructor(
    format: Format,
    engine: EngineReplica,
    saveEngine: () => Promise<Uint8Array>,
    joinPayload: () => JoinPayload,
    notAppliedIsNoOp = false,
  ) {
    this.format = format;
    this.clientId = engine.clientId;
    this.#engine = engine;
    this.#saveEngine = saveEngine;
    this.#joinPayload = joinPayload;
    this.#notAppliedIsNoOp = notAppliedIsNoOp;
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

  applyUpdate(
    update: Uint8Array,
  ): { accepted: true } | { accepted: false; reason: string } {
    this.#assertAlive();
    if (isEmptyUpdate(update)) return { accepted: true };
    let result: unknown;
    try {
      result = this.#engine.applyUpdate(update.slice());
    } catch (error) {
      return rejected(error);
    }
    if (
      result !== null &&
      typeof result === "object" &&
      "applied" in result &&
      (result as { applied: unknown }).applied === false
    ) {
      if (this.#notAppliedIsNoOp) return { accepted: true };
      return { accepted: false, reason: "engine rejected update (not applied)" };
    }
    this.#version += 1;
    return { accepted: true };
  }

  joinPayload(): JoinPayload {
    this.#assertAlive();
    return this.#joinPayload();
  }

  save(): Promise<Uint8Array> {
    this.#assertAlive();
    const run = this.#saveChain.then(() => {
      this.#assertAlive();
      return this.#saveEngine();
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
    this.#engine.dispose();
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
      true,
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
