import { randomUUID } from "node:crypto";
import type { WopiConfig } from "../config";
import { resolveHost } from "../host/allowlist";
import type { HostClient } from "../host/client";
import { FileSession, SessionError, type JoinInput, type JoinResult } from "./fileSession";
import { FileJournal } from "./journal";
import type { SessionLogger, SessionScheduler } from "./scheduler";
import type { openServerReplica } from "../replica/replica";

/** Dependencies for a session registry. */
export interface SessionRegistryDeps {
  config: WopiConfig;
  hostClient: HostClient;
  instanceId?: string;
  journalDir?: string;
  maxEditorsPerFile?: number;
  autosave?: boolean;
  logger?: SessionLogger;
  now?: () => number;
  scheduler?: SessionScheduler;
  openReplica?: typeof openServerReplica;
}

/** Registry of live file sessions keyed by file key. */
export class SessionRegistry {
  readonly instanceId: string;
  #config: WopiConfig;
  #deps: SessionRegistryDeps;
  #sessions = new Map<string, FileSession>();

  constructor(deps: SessionRegistryDeps) {
    this.#deps = deps;
    this.#config = deps.config;
    this.instanceId = deps.instanceId ?? deps.config.instanceId ?? randomUUID();
  }

  /** File key for a host and WOPISrc, without query or fragment. */
  static fileKey(hostId: string, wopiSrc: string): string {
    const url = new URL(wopiSrc);
    url.search = "";
    url.hash = "";
    return `${hostId}:${url.toString()}`;
  }

  /** File key for a WOPISrc under this registry. */
  fileKey(hostId: string, wopiSrc: string): string {
    return SessionRegistry.fileKey(hostId, wopiSrc);
  }

  get size(): number {
    return this.#sessions.size;
  }

  /** Finds a live session by file key. */
  get(fileKey: string): FileSession | null {
    const session = this.#sessions.get(fileKey);
    return session && session.state !== "closed" ? session : null;
  }

  /** Lists live sessions. */
  list(): FileSession[] {
    return [...this.#sessions.values()].filter((session) => session.state !== "closed");
  }

  /** Removes a session from the registry without closing it. */
  remove(fileKey: string): boolean {
    return this.#sessions.delete(fileKey);
  }

  /** Creates a session for a key, replacing any closed entry. */
  create(hostId: string, wopiSrc: string): FileSession {
    const key = SessionRegistry.fileKey(hostId, wopiSrc);
    const existing = this.get(key);
    if (existing) return existing;
    const url = new URL(wopiSrc);
    url.search = "";
    url.hash = "";
    const journalDir = this.#deps.journalDir ?? this.#config.journalDir;
    const session = new FileSession({
      hostClient: this.#deps.hostClient,
      hostId,
      wopiSrc: url.toString(),
      fileKey: key,
      instanceId: this.instanceId,
      logger: this.#deps.logger,
      now: this.#deps.now,
      scheduler: this.#deps.scheduler,
      journal: journalDir ? new FileJournal(journalDir, key) : null,
      maxEditorsPerFile: this.#deps.maxEditorsPerFile ?? this.#config.maxEditorsPerFile ?? 50,
      openReplica: this.#deps.openReplica,
      autosave: this.#deps.autosave,
      onClosed: (closed) => {
        if (this.#sessions.get(closed.fileKey) === closed) this.#sessions.delete(closed.fileKey);
      },
    });
    this.#sessions.set(key, session);
    return session;
  }

  /** Alias for create. */
  getOrCreate(hostId: string, wopiSrc: string): FileSession {
    return this.create(hostId, wopiSrc);
  }

  /** Finds a session by host and WOPISrc. */
  find(hostId: string, wopiSrc: string): FileSession | null {
    return this.get(SessionRegistry.fileKey(hostId, wopiSrc));
  }

  /** Joins a user, creating the file session on first use. */
  async join(input: JoinInput & { hostId?: string }): Promise<JoinResult> {
    const host = input.hostId
      ? this.#config.hosts.find((entry) => entry.id === input.hostId)
      : resolveHost(this.#config, input.wopiSrc);
    if (!host) throw new SessionError("unknownHost", "unknown host id");
    const resolved = resolveHost(this.#config, input.wopiSrc);
    if (resolved.id !== host.id) throw new SessionError("notAllowed", "WOPISrc is not allowed for this host");
    for (let attempt = 0; attempt < 2; attempt += 1) {
      const session = this.create(host.id, input.wopiSrc);
      try {
        return await session.join(input);
      } catch (err) {
        if (err instanceof SessionError && err.kind === "closed" && attempt === 0) {
          if (this.#sessions.get(session.fileKey) === session) this.#sessions.delete(session.fileKey);
          continue;
        }
        if (session.users.length === 0 && session.replica === null && !session.hasPendingJoin) {
          session.abandon();
          this.#sessions.delete(session.fileKey);
        } else if (session.users.length === 0 && session.state === "closed" && !session.hasPendingJoin) {
          this.#sessions.delete(session.fileKey);
        }
        throw err;
      }
    }
    throw new SessionError("closed", "file session is closed");
  }

  /** Closes and removes one session. */
  async close(fileKey: string): Promise<void> {
    const session = this.#sessions.get(fileKey);
    if (!session) return;
    await session.close();
    this.#sessions.delete(fileKey);
  }

  /** Closes every session. */
  async closeAll(): Promise<void> {
    for (const session of [...this.#sessions.values()]) await session.close();
    this.#sessions.clear();
  }
}
