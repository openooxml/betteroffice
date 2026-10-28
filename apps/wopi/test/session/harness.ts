import { generateKeyPair, createProofSigner } from "../../src/proof";
import { createHostClient, type HostFetch } from "../../src/host/client";
import { createReferenceHost, type ReferenceHost } from "../host/reference-host";
import { sourceBytes } from "../replica/clients";
import { SessionRegistry } from "../../src/session/registry";
import type { SessionLogger, SessionScheduler } from "../../src/session/scheduler";
import type { Format } from "../../src/replica/replica";

export const START = 1700000000000;
export const SRC = "https://cloud.example.com/wopi/files/file1";

interface Timer {
  id: number;
  at: number;
  fn: () => void;
  interval: number | null;
}

export class FakeClock {
  t = START;
  #timers = new Map<number, Timer>();
  #next = 1;
  now = (): number => this.t;
  scheduler: SessionScheduler = {
    setTimeout: (fn, ms) => {
      const id = this.#next++;
      this.#timers.set(id, { id, at: this.t + Math.max(0, ms), fn, interval: null });
      return id as unknown as ReturnType<typeof setTimeout>;
    },
    clearTimeout: (handle) => {
      this.#timers.delete(handle as unknown as number);
    },
    setInterval: (fn, ms) => {
      const id = this.#next++;
      this.#timers.set(id, { id, at: this.t + Math.max(1, ms), fn, interval: Math.max(1, ms) });
      return id as unknown as ReturnType<typeof setInterval>;
    },
    clearInterval: (handle) => {
      this.#timers.delete(handle as unknown as number);
    },
  };

  async flush(): Promise<void> {
    for (let i = 0; i < 20; i += 1) await new Promise((r) => setImmediate(r));
  }

  async drain(rounds = 100): Promise<void> {
    for (let i = 0; i < rounds; i += 1) {
      await new Promise((r) => setTimeout(r, 0));
      await new Promise((r) => setImmediate(r));
    }
  }

  async advance(ms: number): Promise<void> {
    const target = this.t + ms;
    for (;;) {
      let next: Timer | undefined;
      for (const timer of this.#timers.values()) {
        if (timer.at <= target && (!next || timer.at < next.at || (timer.at === next.at && timer.id < next.id))) next = timer;
      }
      if (!next) break;
      this.t = next.at;
      if (next.interval !== null) next.at = this.t + next.interval;
      else this.#timers.delete(next.id);
      next.fn();
      await this.flush();
    }
    this.t = target;
    await this.flush();
    await this.drain(12);
  }
}

export interface RecordedCall {
  token: string | null;
  override: string | null;
  pathname: string;
  headers: Headers;
}

export interface Harness {
  clock: FakeClock;
  host: ReferenceHost;
  hostClient: ReturnType<typeof createHostClient>;
  registry: SessionRegistry;
  calls: RecordedCall[];
  logs: string[];
  logger: SessionLogger;
  config: ConstructorParameters<typeof SessionRegistry>[0]["config"];
  scheduler: SessionScheduler;
  timerDelays: number[];
  putCount(): number;
  callsFor(override: string): RecordedCall[];
}

export function makeHarness(options: {
  journalDir?: string;
  maxEditorsPerFile?: number;
  instanceId?: string;
  quirks?: { timestampConflict?: boolean; relativeUrlToken?: string };
  format?: Format;
  fileName?: string;
  lockFails?: boolean;
  failFirstRead?: boolean;
  beforePut?: () => Promise<void>;
  beforeUnlock?: () => Promise<void>;
  beforeCheckFileInfo?: () => Promise<void>;
  refreshStatus?: () => number | null;
  lockStatus?: () => number | null;
  getLockStatus?: () => number | null;
  relativeStatus?: () => number | null;
} = {}): Harness {
  const clock = new FakeClock();
  const pair = generateKeyPair();
  const signer = createProofSigner(pair);
  const host = createReferenceHost({ now: clock.now, quirks: options.quirks });
  const format = options.format ?? "xlsx";
  host.setFile("file1", sourceBytes(format), { name: options.fileName ?? `book.${format}` });
  const calls: RecordedCall[] = [];
  let reads = 0;
  const fetchFn: HostFetch = async (input, init) => {
    const request = new Request(input as string, init as RequestInit);
    const url = new URL(request.url);
    const override = request.headers.get("X-WOPI-Override");
    calls.push({ token: url.searchParams.get("access_token"), override, pathname: url.pathname, headers: request.headers });
    if (options.lockFails && override === "LOCK" && !request.headers.get("X-WOPI-OldLock")) {
      return new Response(JSON.stringify({ error: "lock boom" }), { status: 500, headers: { "Content-Type": "application/json" } });
    }
    if (options.failFirstRead && request.method === "GET" && url.pathname.endsWith("/contents")) {
      reads += 1;
      if (reads === 1) return new Response(JSON.stringify({ error: "read boom" }), { status: 500, headers: { "Content-Type": "application/json" } });
    }
    if (override === "PUT" && options.beforePut) await options.beforePut();
    if (override === "UNLOCK" && options.beforeUnlock) await options.beforeUnlock();
    if (request.method === "GET" && !url.pathname.endsWith("/contents") && options.beforeCheckFileInfo) await options.beforeCheckFileInfo();
    const status = override === "REFRESH_LOCK" ? options.refreshStatus?.()
      : override === "PUT_RELATIVE" ? options.relativeStatus?.()
      : override === "LOCK" && !request.headers.get("X-WOPI-OldLock") ? options.lockStatus?.()
      : override === "GET_LOCK" ? options.getLockStatus?.()
      : null;
    if (status) return new Response(null, { status });
    return host.fetch(request);
  };
  const client = createHostClient({ signer, maxFileBytes: 64 * 1024 * 1024, fetch: fetchFn, now: clock.now });
  const logs: string[] = [];
  const logger: SessionLogger = {
    info: (message, meta) => logs.push(`info ${message} ${meta ? JSON.stringify(meta) : ""}`),
    warn: (message, meta) => logs.push(`warn ${message} ${meta ? JSON.stringify(meta) : ""}`),
    error: (message, meta) => logs.push(`error ${message} ${meta ? JSON.stringify(meta) : ""}`),
  };
  const timerDelays: number[] = [];
  const scheduler: SessionScheduler = {
    setTimeout: (fn, ms) => {
      timerDelays.push(ms);
      return clock.scheduler.setTimeout(fn, ms);
    },
    clearTimeout: (handle) => clock.scheduler.clearTimeout(handle),
    setInterval: (fn, ms) => {
      timerDelays.push(ms);
      return clock.scheduler.setInterval(fn, ms);
    },
    clearInterval: (handle) => clock.scheduler.clearInterval(handle),
  };
  const config: Harness["config"] = {
    publicUrl: "https://office.example.com",
    proofKeyFile: "data/proof-key.pem",
    maxFileBytes: 64 * 1024 * 1024,
    hosts: [{ id: "testhost", wopiSrcOrigins: ["https://cloud.example.com"], frameOrigins: ["https://cloud.example.com"] }],
    journalDir: options.journalDir,
    instanceId: options.instanceId ?? "inst-test",
    maxEditorsPerFile: options.maxEditorsPerFile ?? 50,
  };
  const registry = new SessionRegistry({
    config,
    hostClient: client,
    logger,
    now: clock.now,
    scheduler,
  });
  return {
    clock,
    host,
    hostClient: client,
    registry,
    config,
    scheduler,
    timerDelays,
    calls,
    logs,
    logger,
    putCount: () => calls.filter((c) => c.override === "PUT").length,
    callsFor: (override) => calls.filter((c) => c.override === override),
  };
}

export function addToken(h: Harness, token: string, userId: string, opts: { canWrite?: boolean; expiresAt?: number; name?: string } = {}): void {
  h.host.tokens.set(token, { userId, userName: opts.name ?? userId, canWrite: opts.canWrite ?? true, expiresAt: opts.expiresAt });
}

export function ttl(clock: FakeClock, msFromNow: number): number {
  return clock.t + msFromNow;
}
