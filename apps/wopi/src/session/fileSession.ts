import { randomInt, randomUUID } from "node:crypto";
import type { HostClient } from "../host/client";
import { WopiHostError } from "../host/errors";
import { redactText } from "../host/redact";
import { openServerReplica, type Format, type JoinPayload, type ServerReplica } from "../replica/replica";
import { electLeader } from "./leader";
import { UserSession } from "./userSession";
import { FileJournal } from "./journal";
import { noopLogger, systemScheduler, type SessionLogger, type SessionScheduler } from "./scheduler";

export const LOCK_REFRESH_MS = 10 * 60 * 1000;
export const AUTOSAVE_DEBOUNCE_MS = 30 * 1000;
export const AUTOSAVE_MAX_MS = 5 * 60 * 1000;
export const HEARTBEAT_TIMEOUT_MS = 60 * 1000;
export const HOLD_FLUSH_MS = 30 * 60 * 1000;
export const TOKEN_EXPIRY_WARNING_MS = 5 * 60 * 1000;
export const OUR_LOCK_PREFIX = "bo1:";
export const MAX_TIMER_MS = 2147483647;

export type FileSessionState = "opening" | "open" | "closing" | "closed";
export type ViewOnlyReason = "lockedElsewhere";

/** Failure raised by the session layer. */
export class SessionError extends Error {
  readonly kind: string;
  constructor(kind: string, message: string) {
    super(message);
    this.name = "SessionError";
    this.kind = kind;
  }
}

/** Events emitted by a file session. Payloads never carry tokens or content. */
export interface FileSessionEvents {
  leaderChanged: { leaderId: string | null; previousLeaderId: string | null };
  saved: { itemVersion: string | null; by: string | null };
  saveFailed: { kind: string };
  conflict: { name: string; url: string; copyName: string; copyUrl: string };
  viewOnly: { reason: ViewOnlyReason };
  replicaUpdate: { update: Uint8Array };
  closed: Record<string, never>;
  lostFlush: Record<string, never>;
  tokenExpiring: { userId: string; hostUserId: string; expiresAt: number | null };
}

export type FileSessionEventName = keyof FileSessionEvents;
type Listener<T> = (payload: T) => void;

/** Join request for one user. */
export interface JoinInput {
  wopiSrc: string;
  accessToken: string;
  accessTokenTtl: number;
  mode: "edit" | "view";
}

/** File summary returned on join. */
export interface JoinFileInfo {
  name: string;
  size: number;
  canWrite: boolean;
  readOnly: boolean;
}

/** Result of joining a file session. */
export interface JoinResult {
  user: UserSession;
  file: JoinFileInfo;
  joinPayload: JoinPayload;
  viewOnlyReason?: ViewOnlyReason;
  session: FileSession;
  sessionKey: string;
}

/** Result of an explicit or automatic save attempt. */
export interface SaveResult {
  saved: boolean;
  itemVersion: string | null;
  by: string | null;
  conflictName?: string;
  conflictUrl?: string;
}

/** Dependencies for one file session. */
export interface FileSessionDeps {
  hostClient: HostClient;
  hostId: string;
  wopiSrc: string;
  fileKey: string;
  instanceId: string;
  logger?: SessionLogger;
  now?: () => number;
  scheduler?: SessionScheduler;
  journal?: FileJournal | null;
  maxEditorsPerFile?: number;
  openReplica?: typeof openServerReplica;
  onClosed?: (session: FileSession) => void;
}

interface SaveOptions {
  tokenUser?: UserSession | null;
  isAutosave: boolean;
  isExitSave: boolean;
  by?: string | null;
}

function formatFromName(name: string): Format {
  const ext = name.slice(name.lastIndexOf(".") + 1).toLowerCase();
  if (ext === "docx" || ext === "xlsx" || ext === "pptx") return ext;
  throw new SessionError("unsupportedFormat", `unsupported file extension: ${ext}`);
}

function conflictTargetName(name: string, now: number): string {
  const date = new Date(now);
  const pad = (value: number) => String(value).padStart(2, "0");
  const stamp = `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())} ${pad(date.getUTCHours())}.${pad(date.getUTCMinutes())}`;
  const dot = name.lastIndexOf(".");
  if (dot > 0) return `${name.slice(0, dot)} (conflict ${stamp}).${name.slice(dot + 1)}`;
  return `${name} (conflict ${stamp})`;
}

/** One file session: users, lock, replica, leader and saves for one file. */
export class FileSession {
  readonly hostId: string;
  readonly wopiSrc: string;
  readonly fileKey: string;
  readonly lockId: string;
  #deps: FileSessionDeps;
  #host: HostClient;
  #logger: SessionLogger;
  #now: () => number;
  #scheduler: SessionScheduler;
  #state: FileSessionState = "opening";
  #users = new Map<string, UserSession>();
  #history: UserSession[] = [];
  #replica: ServerReplica | null = null;
  #leader: UserSession | null = null;
  #listeners = new Map<string, Set<Listener<never>>>();
  #opChain: Promise<unknown> = Promise.resolve();
  #fileName = "";
  #fileSize = 0;
  #lastModifiedTime: string | null = null;
  #itemVersion: string | null = null;
  #lockHeld = false;
  #lastLockRefreshAt = 0;
  #viewOnlyReason: ViewOnlyReason | null = null;
  #savedVersion = 0;
  #firstDirtyAt: number | null = null;
  #lastUpdateAt = 0;
  #everHadEditor = false;
  #holdPending = false;
  #refreshTimer: ReturnType<typeof setTimeout> | null = null;
  #autosaveTimer: ReturnType<typeof setTimeout> | null = null;
  #holdTimer: ReturnType<typeof setTimeout> | null = null;
  #sweepTimer: ReturnType<typeof setInterval> | null = null;
  #userTimers = new Map<string, Array<ReturnType<typeof setTimeout>>>();
  #format: Format | null = null;
  #journalActive = false;
  #journalVersion: string | null = null;
  #pendingJoins = 0;
  #sourceBytes: Uint8Array | null = null;
  #offLocalUpdate: (() => void) | null = null;
  #openedClientId: number | null = null;
  #openedInitialState: Uint8Array | null = null;
  #replayStats: { accepted: number; rejected: number } | null = null;

  constructor(deps: FileSessionDeps) {
    this.#deps = deps;
    this.hostId = deps.hostId;
    this.wopiSrc = deps.wopiSrc;
    this.fileKey = deps.fileKey;
    this.lockId = `${OUR_LOCK_PREFIX}${deps.instanceId}:${randomUUID()}`;
    this.#host = deps.hostClient;
    this.#logger = deps.logger ?? noopLogger;
    this.#now = deps.now ?? (() => Date.now());
    this.#scheduler = deps.scheduler ?? systemScheduler;
    this.#sweepTimer = this.#scheduler.setInterval(() => {
      this.#enqueueTask(() => this.#sweepInternal());
    }, 5000);
  }

  get state(): FileSessionState {
    return this.#state;
  }

  get users(): UserSession[] {
    return [...this.#users.values()];
  }

  get leader(): UserSession | null {
    return this.#leader;
  }

  get replica(): ServerReplica | null {
    return this.#replica;
  }

  get format(): Format | null {
    return this.#format;
  }

  get lockHeld(): boolean {
    return this.#lockHeld;
  }

  get viewOnly(): boolean {
    return this.#viewOnlyReason !== null;
  }

  get viewOnlyReason(): ViewOnlyReason | null {
    return this.#viewOnlyReason;
  }

  get dirty(): boolean {
    return this.#replica !== null && this.#replica.version > this.#savedVersion;
  }

  get itemVersion(): string | null {
    return this.#itemVersion;
  }

  get awaitingFlushToken(): boolean {
    return this.#holdPending;
  }

  /** Accepted and rejected counts from the last journal replay. */
  get replayStats(): { accepted: number; rejected: number } | null {
    return this.#replayStats;
  }

  /** Subscribes to a session event and returns an unsubscribe function. */
  on<E extends FileSessionEventName>(event: E, listener: Listener<FileSessionEvents[E]>): () => void {
    let set = this.#listeners.get(event);
    if (!set) {
      set = new Set();
      this.#listeners.set(event, set);
    }
    set.add(listener as Listener<never>);
    return () => {
      set.delete(listener as Listener<never>);
    };
  }

  /** Removes a session event listener. */
  off<E extends FileSessionEventName>(event: E, listener: Listener<FileSessionEvents[E]>): void {
    this.#listeners.get(event)?.delete(listener as Listener<never>);
  }

  #emit<E extends FileSessionEventName>(event: E, payload: FileSessionEvents[E]): void {
    const set = this.#listeners.get(event);
    if (!set) return;
    for (const listener of [...set]) {
      try {
        (listener as Listener<FileSessionEvents[E]>)(payload);
      } catch {}
    }
  }

  #scrub(text: string): string {
    let out = redactText(text);
    for (const user of this.#history) {
      if (user.accessToken) out = out.split(user.accessToken).join("***");
    }
    return out;
  }

  #log(level: "info" | "warn" | "error", message: string, meta?: Record<string, unknown>): void {
    const safe = this.#scrub(message);
    if (level === "info") this.#logger.info(safe, meta);
    else if (level === "warn") this.#logger.warn(safe, meta);
    else this.#logger.error(safe, meta);
  }

  /** Returns a user by session handle id. */
  getUser(id: string): UserSession | null {
    return this.#users.get(id) ?? null;
  }

  /** Whether a join call is queued or running. */
  get hasPendingJoin(): boolean {
    return this.#pendingJoins > 0;
  }

  #enqueue<T>(task: () => Promise<T> | T): Promise<T> {
    const run = this.#opChain.then(task);
    this.#opChain = run.then(
      () => undefined,
      () => undefined,
    );
    return run;
  }

  #enqueueTask(task: () => Promise<unknown> | unknown): void {
    void this.#enqueue(task).catch(() => {
      this.#log("warn", "session task failed");
    });
  }

  /** Joins a user to this file session. */
  join(input: JoinInput): Promise<JoinResult> {
    this.#pendingJoins += 1;
    return this.#enqueue(() => this.#joinInternal(input)).finally(() => {
      this.#pendingJoins -= 1;
    });
  }

  async #joinInternal(input: JoinInput): Promise<JoinResult> {
    if (this.#state === "closed" || this.#state === "closing") {
      throw new SessionError("closed", "file session is closed");
    }
    const max = this.#deps.maxEditorsPerFile ?? 50;
    const connectedCount = [...this.#users.values()].filter((user) => user.connected).length;
    if (connectedCount >= max) {
      throw new SessionError("tooManyEditors", "max editors per file reached");
    }
    const info = await this.#host.checkFileInfo({ wopiSrc: this.wopiSrc, accessToken: input.accessToken });
    const wantsWrite = input.mode === "edit" && info.UserCanWrite === true && info.ReadOnly !== true;
    const first = this.#replica === null;
    if (first) {
      try {
        this.#state = "opening";
        this.#format = formatFromName(info.BaseFileName);
        this.#fileName = info.BaseFileName;
        this.#fileSize = info.Size;
        this.#lastModifiedTime = info.LastModifiedTime ?? null;
        this.#itemVersion = info.Version ?? null;
        const bytes = await this.#host.readFile({ wopiSrc: this.wopiSrc, accessToken: input.accessToken });
        this.#sourceBytes = bytes;
        const open = this.#deps.openReplica ?? openServerReplica;
        this.#attachReplica(await open(this.#format, bytes, { clientId: randomInt(1, 2000000000) }));
        this.#openedClientId = this.#replica!.clientId;
        try {
          this.#openedInitialState = this.#replica!.encodeStateAsUpdate();
        } catch {
          this.#openedInitialState = null;
        }
        this.#journalVersion = info.Version ?? null;
        this.#state = "open";
      } catch (err) {
        this.#failOpening();
        throw err;
      }
    } else {
      this.#fileName = info.BaseFileName || this.#fileName;
      this.#fileSize = info.Size;
    }
    const lockHeldBefore = this.#lockHeld;
    if (!this.#lockHeld && this.#viewOnlyReason === null && wantsWrite) {
      try {
        await this.#attemptLock(input.accessToken);
      } catch (err) {
        if (first) this.#failOpening();
        throw err;
      }
    }
    const lockJustAcquired = !lockHeldBefore && this.#lockHeld;
    if (this.#lockHeld && !this.#journalActive) {
      try {
        await this.#activateJournal(info.Version ?? this.#journalVersion ?? this.#itemVersion);
      } catch (err) {
        if (first) this.#failOpening();
        throw err;
      }
    }
    const canWrite = wantsWrite && this.#viewOnlyReason === null;
    const now = this.#now();
    const user = new UserSession({
      hostId: this.hostId,
      wopiSrc: this.wopiSrc,
      accessToken: input.accessToken,
      accessTokenTtl: input.accessTokenTtl,
      userId: info.UserId,
      friendlyName: info.UserFriendlyName ?? info.UserId,
      canWrite,
      joinedAt: now,
    });
    this.#users.set(user.id, user);
    this.#history.push(user);
    if (canWrite) this.#everHadEditor = true;
    this.#scheduleUserTimers(user);
    await this.#refreshLeader(lockJustAcquired);
    if (canWrite && this.dirty && !this.#holdPending && this.#autosaveTimer === null) this.#scheduleAutosave();
    if (this.#holdPending && canWrite) {
      this.#holdPending = false;
      this.#clearTimer("hold");
      if (this.dirty) {
        await this.#saveInternal({ tokenUser: user, isAutosave: false, isExitSave: false, by: user.id });
      }
      if (this.dirty) this.#scheduleAutosave();
    }
    const result: JoinResult = {
      user,
      file: { name: this.#fileName, size: this.#fileSize, canWrite, readOnly: !canWrite },
      joinPayload: this.#replica!.joinPayload(),
      session: this,
      sessionKey: this.fileKey,
    };
    if (this.#viewOnlyReason) result.viewOnlyReason = this.#viewOnlyReason;
    return result;
  }

  #attachReplica(replica: ServerReplica): void {
    this.#replica = replica;
    this.#offLocalUpdate = replica.onLocalUpdate((update) => {
      try {
        if (this.#journalActive) this.#deps.journal?.appendUpdate(update);
      } catch {
        this.#log("warn", "journal append failed");
      }
      this.#emit("replicaUpdate", { update });
    });
  }

  #destroyReplica(): void {
    if (this.#offLocalUpdate) {
      try {
        this.#offLocalUpdate();
      } catch {}
      this.#offLocalUpdate = null;
    }
    if (this.#replica) {
      try {
        this.#replica.destroy();
      } catch {}
      this.#replica = null;
    }
  }

  #failOpening(): void {
    this.#state = "opening";
    this.#clearTimer("refresh");
    this.#clearTimer("autosave");
    this.#clearTimer("hold");
    this.#destroyReplica();
    this.#lockHeld = false;
    this.#viewOnlyReason = null;
    this.#leader = null;
    this.#journalActive = false;
    this.#deps.journal?.close();
  }

  async #activateJournal(currentVersion: string | null): Promise<void> {
    const journal = this.#deps.journal;
    if (!journal || !this.#replica || !this.#format) return;
    this.#journalVersion = currentVersion;
    const existing = journal.read();
    const hasJournal = existing !== null && existing.source !== null && existing.itemVersion !== null;
    if (hasJournal) {
      const known = currentVersion !== null && (existing.itemVersion === currentVersion || existing.writtenVersions.includes(currentVersion));
      if (known) {
        if (this.#replica.version === 0 && !this.dirty) {
          const source = existing.source!;
          this.#destroyReplica();
          const open = this.#deps.openReplica ?? openServerReplica;
          this.#attachReplica(await open(this.#format, source, { clientId: existing.clientId ?? randomInt(1, 2000000000) }));
          this.#savedVersion = 0;
          const versionBefore = this.#replica.version;
          let accepted = 0;
          let rejected = 0;
          if (existing.initialState) {
            const initial = this.#replica.applyUpdate(existing.initialState);
            if (initial.accepted) accepted += 1;
            else rejected += 1;
          }
          for (const update of existing.updates) {
            const result = this.#replica.applyUpdate(update);
            if (result.accepted) accepted += 1;
            else {
              rejected += 1;
              this.#log("warn", "journal replay update rejected");
            }
          }
          this.#replayStats = { accepted, rejected };
          if (this.#replica.version !== versionBefore || existing.updates.length > 0) {
            this.#firstDirtyAt = this.#now();
            this.#lastUpdateAt = this.#now();
          }
          if (existing.updates.length > 0) this.#log("info", "journal replayed", { updates: existing.updates.length });
        }
        this.#journalActive = true;
        journal.begin(existing.source!, existing.itemVersion, existing.clientId, existing.initialState);
        return;
      }
      this.#log("warn", "journal discarded: host version changed");
      journal.delete();
    }
    this.#journalActive = true;
    if (this.#sourceBytes) journal.begin(this.#sourceBytes, currentVersion, this.#openedClientId, this.#openedInitialState);
  }

  async #attemptLock(accessToken: string): Promise<void> {
    try {
      await this.#host.lock({ wopiSrc: this.wopiSrc, accessToken, lock: this.lockId });
      this.#lockHeld = true;
      this.#lastLockRefreshAt = this.#now();
      this.#scheduleRefresh();
      return;
    } catch (err) {
      if (!(err instanceof WopiHostError) || (err.kind !== "lockMismatch" && err.kind !== "conflict")) throw err;
      let current = err.currentLock ?? null;
      if (current === null) {
        try {
          current = (await this.#host.getLock({ wopiSrc: this.wopiSrc, accessToken })).lock;
        } catch {}
      }
      if (current && current.startsWith(OUR_LOCK_PREFIX)) {
        try {
          await this.#host.unlockAndRelock({ wopiSrc: this.wopiSrc, accessToken, lock: this.lockId, oldLock: current });
          this.#lockHeld = true;
          this.#lastLockRefreshAt = this.#now();
          this.#scheduleRefresh();
          return;
        } catch (relockErr) {
          if (!(relockErr instanceof WopiHostError) || (relockErr.kind !== "lockMismatch" && relockErr.kind !== "conflict")) throw relockErr;
        }
      }
      this.#viewOnlyReason = "lockedElsewhere";
      this.#log("warn", "file is locked elsewhere, opening view only");
      this.#emit("viewOnly", { reason: "lockedElsewhere" });
    }
  }

  #clearTimer(which: "refresh" | "autosave" | "hold"): void {
    if (which === "refresh" && this.#refreshTimer !== null) {
      this.#scheduler.clearTimeout(this.#refreshTimer);
      this.#refreshTimer = null;
    }
    if (which === "autosave" && this.#autosaveTimer !== null) {
      this.#scheduler.clearTimeout(this.#autosaveTimer);
      this.#autosaveTimer = null;
    }
    if (which === "hold" && this.#holdTimer !== null) {
      this.#scheduler.clearTimeout(this.#holdTimer);
      this.#holdTimer = null;
    }
  }

  #clearUserTimers(id: string): void {
    const handles = this.#userTimers.get(id);
    if (handles) for (const handle of handles) this.#scheduler.clearTimeout(handle);
    this.#userTimers.delete(id);
  }

  #clearAllTimers(): void {
    this.#clearTimer("refresh");
    this.#clearTimer("autosave");
    this.#clearTimer("hold");
    if (this.#sweepTimer !== null) {
      this.#scheduler.clearInterval(this.#sweepTimer);
      this.#sweepTimer = null;
    }
    for (const id of [...this.#userTimers.keys()]) this.#clearUserTimers(id);
  }

  #setTimerAt(target: number, fn: () => void, handles: Array<ReturnType<typeof setTimeout>>): void {
    const delay = Math.max(0, target - this.#now());
    const step = Math.min(delay, MAX_TIMER_MS);
    const handle = this.#scheduler.setTimeout(() => {
      if (this.#now() < target) {
        this.#setTimerAt(target, fn, handles);
        return;
      }
      fn();
    }, step);
    handles.push(handle);
  }

  #scheduleUserTimers(user: UserSession): void {
    if (user.expiresAt === null) return;
    const handles: Array<ReturnType<typeof setTimeout>> = [];
    this.#userTimers.set(user.id, handles);
    this.#setTimerAt(user.expiresAt - TOKEN_EXPIRY_WARNING_MS, () => {
      if (this.#users.has(user.id)) {
        this.#emit("tokenExpiring", { userId: user.id, hostUserId: user.userId, expiresAt: user.expiresAt });
      }
    }, handles);
    this.#setTimerAt(user.expiresAt - 2 * 60 * 1000 + 1, () => {
      this.#enqueueTask(async () => {
        if (this.#users.has(user.id)) await this.#refreshLeader();
      });
    }, handles);
    this.#setTimerAt(user.expiresAt, () => {
      this.#enqueueTask(async () => {
        if (!this.#users.has(user.id) || this.#state !== "open") return;
        user.canWrite = false;
        await this.#refreshLeader();
        await this.#maybeCloseAfterDeparture([user]);
      });
    }, handles);
  }

  #syncLeader(): UserSession | null {
    const previous = this.#leader;
    const next = electLeader([...this.#users.values()], this.#now());
    this.#leader = next;
    if (previous?.id !== next?.id) {
      this.#emit("leaderChanged", { leaderId: next?.id ?? null, previousLeaderId: previous?.id ?? null });
    }
    return next;
  }

  async #refreshLeader(suppressRefresh = false): Promise<UserSession | null> {
    const previous = this.#leader;
    const next = this.#syncLeader();
    if (previous?.id !== next?.id) {
      if (this.dirty) this.#scheduleAutosave();
      if (next && this.#lockHeld && !suppressRefresh) {
        await this.#refreshLockNow();
        this.#scheduleRefresh();
      }
    }
    return next;
  }

  #scheduleRefresh(): void {
    this.#clearTimer("refresh");
    if (!this.#lockHeld || this.#state !== "open") return;
    this.#refreshTimer = this.#scheduler.setTimeout(() => {
      this.#enqueueTask(async () => {
        if (this.#state !== "open" || !this.#lockHeld) return;
        await this.#refreshLockNow();
        this.#scheduleRefresh();
      });
    }, LOCK_REFRESH_MS);
  }

  async #refreshLockNow(): Promise<void> {
    if (!this.#lockHeld) return;
    const tried = new Set<string>();
    let tokenUser: UserSession | null = this.#leader && this.#leader.isTokenValidAt(this.#now()) ? this.#leader : this.#pickToken([]);
    while (tokenUser) {
      try {
        await this.#host.refreshLock({ wopiSrc: this.wopiSrc, accessToken: tokenUser.accessToken, lock: this.lockId });
        this.#lastLockRefreshAt = this.#now();
        return;
      } catch (err) {
        const kind = err instanceof WopiHostError ? err.kind : "unknown";
        this.#log("warn", `refresh lock failed (${kind})`);
        if (kind !== "unauthorized") return;
        tried.add(tokenUser.id);
        this.#markTokenDead(tokenUser);
        tokenUser = this.#pickToken([], tried);
      }
    }
  }

  #pickToken(preferred: UserSession[], excluded?: Set<string>): UserSession | null {
    const now = this.#now();
    const seen = new Set<string>();
    const candidates: UserSession[] = [];
    for (const user of [...preferred, ...this.#history]) {
      if (seen.has(user.id)) continue;
      seen.add(user.id);
      candidates.push(user);
    }
    for (const user of candidates) {
      if (excluded?.has(user.id)) continue;
      if (user.canWrite && user.accessToken && user.isTokenValidAt(now)) return user;
    }
    return null;
  }

  #isConnectedEditor(user: UserSession): boolean {
    return user.connected && user.canWrite && this.#viewOnlyReason === null && user.isTokenValidAt(this.#now());
  }

  #connectedEditors(): UserSession[] {
    return [...this.#users.values()].filter((user) => this.#isConnectedEditor(user));
  }

  #markTokenDead(user: UserSession): void {
    if (user.expiresAt === null || user.expiresAt > this.#now()) user.expiresAt = this.#now();
    user.canWrite = false;
    this.#syncLeader();
  }

  /** Applies one update from a connected writer. */
  applyUpdate(userId: string, update: Uint8Array): { accepted: true } | { accepted: false; reason: string } {
    const user = this.#users.get(userId);
    if (!user) throw new SessionError("unknownUser", "unknown user");
    if (this.#state !== "open" || !this.#replica) return { accepted: false, reason: "session is not open" };
    if (!user.connected) return { accepted: false, reason: "user is not connected" };
    if (!user.canWrite || this.#viewOnlyReason !== null) return { accepted: false, reason: "user is read only" };
    if (!user.isTokenValidAt(this.#now())) return { accepted: false, reason: "token expired" };
    const before = this.#replica.version;
    const result = this.#replica.applyUpdate(update);
    if (!result.accepted) return result;
    user.lastSeen = this.#now();
    if (this.#replica.version !== before) {
      if (this.#firstDirtyAt === null) this.#firstDirtyAt = this.#now();
      this.#lastUpdateAt = this.#now();
      try {
        if (this.#journalActive) this.#deps.journal?.appendUpdate(update);
      } catch {
        this.#log("warn", "journal append failed");
      }
      this.#scheduleAutosave();
    }
    return result;
  }

  #scheduleAutosave(): void {
    this.#clearTimer("autosave");
    if (!this.dirty || this.#state !== "open") return;
    const first = this.#firstDirtyAt ?? this.#lastUpdateAt;
    const target = Math.min(this.#lastUpdateAt + AUTOSAVE_DEBOUNCE_MS, first + AUTOSAVE_MAX_MS);
    const delay = Math.max(0, target - this.#now());
    this.#autosaveTimer = this.#scheduler.setTimeout(() => {
      this.#enqueueTask(() => this.#runAutosave());
    }, delay);
  }

  #scheduleAutosaveRetry(): void {
    this.#clearTimer("autosave");
    if (!this.dirty || this.#state !== "open") return;
    this.#autosaveTimer = this.#scheduler.setTimeout(() => {
      this.#enqueueTask(() => this.#runAutosave());
    }, AUTOSAVE_DEBOUNCE_MS);
  }

  async #runAutosave(): Promise<void> {
    if (this.#state !== "open" || !this.dirty) return;
    const versionBefore = this.#replica?.version ?? 0;
    const tokenUser = this.#leader ?? (await this.#refreshLeader()) ?? this.#pickToken([]);
    if (!tokenUser) {
      this.#scheduleAutosaveRetry();
      return;
    }
    const result = await this.#saveInternal({ tokenUser, isAutosave: true, isExitSave: false, by: tokenUser.id });
    if (this.dirty && this.#state === "open") {
      if (result.saved && (this.#replica?.version ?? 0) !== versionBefore) this.#scheduleAutosave();
      else if (result.saved) this.#scheduleAutosaveRetry();
      else this.#scheduleAutosaveRetry();
    }
  }

  /** Marks a user as alive and reconnects them when needed. */
  heartbeat(userId: string): void {
    const user = this.#users.get(userId);
    if (!user) throw new SessionError("unknownUser", "unknown user");
    user.lastSeen = this.#now();
    if (user.connected) return;
    user.connected = true;
    this.#enqueueTask(async () => {
      if (this.#state !== "open" || !this.#users.has(user.id) || !user.connected) return;
      if (this.#holdPending && this.#isConnectedEditor(user)) {
        this.#holdPending = false;
        this.#clearTimer("hold");
      }
      await this.#refreshLeader();
      if (this.dirty) this.#scheduleAutosave();
    });
  }

  /** Marks a user disconnected and runs the departure policy. */
  disconnect(userId: string): Promise<void> {
    return this.#enqueue(() => this.#disconnectInternal(userId));
  }

  async #disconnectInternal(userId: string): Promise<void> {
    const user = this.#users.get(userId);
    if (!user) throw new SessionError("unknownUser", "unknown user");
    if (!user.connected) return;
    user.connected = false;
    await this.#refreshLeader();
    await this.#maybeCloseAfterDeparture([user]);
  }

  /** Removes a user and runs the departure policy. */
  leave(userId: string): Promise<void> {
    return this.#enqueue(() => this.#leaveInternal(userId));
  }

  async #leaveInternal(userId: string): Promise<void> {
    const user = this.#users.get(userId);
    if (!user) throw new SessionError("unknownUser", "unknown user");
    const leaderBefore = this.#leader;
    this.#users.delete(userId);
    this.#clearUserTimers(userId);
    user.connected = false;
    await this.#refreshLeader();
    const departing = leaderBefore && leaderBefore.id !== userId ? [leaderBefore, user] : [user];
    await this.#maybeCloseAfterDeparture(departing);
  }

  async #sweepInternal(): Promise<void> {
    if (this.#state !== "open") return;
    const now = this.#now();
    const stale = [...this.#users.values()].filter((user) => user.connected && now - user.lastSeen >= HEARTBEAT_TIMEOUT_MS);
    for (const user of stale) {
      user.connected = false;
      this.#log("info", "user disconnected after heartbeat timeout", { userId: user.id });
    }
    if (stale.length > 0) {
      await this.#refreshLeader();
      await this.#maybeCloseAfterDeparture(stale);
    }
  }

  /** Saves now, crediting the requesting user when their token can write. */
  requestSave(userId: string): Promise<SaveResult> {
    return this.#enqueue(() => this.#requestSaveInternal(userId));
  }

  async #requestSaveInternal(userId: string): Promise<SaveResult> {
    const user = this.#users.get(userId);
    if (!user) throw new SessionError("unknownUser", "unknown user");
    if (this.#state !== "open") throw new SessionError("closed", "file session is not open");
    if (!this.dirty) return { saved: false, itemVersion: this.#itemVersion, by: null };
    const tokenUser = user.canWrite && user.isTokenValidAt(this.#now()) ? user : this.#leader ?? this.#pickToken([]);
    return this.#saveInternal({ tokenUser, isAutosave: false, isExitSave: false, by: tokenUser?.id ?? null });
  }

  #recordWritten(): void {
    const journal = this.#deps.journal;
    if (!journal || !this.#journalActive) return;
    try {
      journal.recordWrittenVersion(this.#itemVersion);
      journal.flush();
    } catch {
      this.#log("warn", "journal write failed");
    }
  }

  async #saveInternal(options: SaveOptions): Promise<SaveResult> {
    if (!this.#replica || this.#state === "closed" || this.#state === "closing") {
      return { saved: false, itemVersion: this.#itemVersion, by: options.by ?? null };
    }
    if (!this.dirty) return { saved: false, itemVersion: this.#itemVersion, by: options.by ?? null };
    const versionAtStart = this.#replica.version;
    let bytes: Uint8Array;
    try {
      bytes = (await this.#replica.save()).bytes;
    } catch {
      this.#emit("saveFailed", { kind: "invalidResponse" });
      this.#log("error", "save failed (invalidResponse)");
      return { saved: false, itemVersion: this.#itemVersion, by: options.by ?? null };
    }
    const headers: Record<string, string> = {
      "X-COOL-WOPI-IsModifiedByUser": "true",
      "X-COOL-WOPI-IsAutosave": options.isAutosave ? "true" : "false",
      "X-COOL-WOPI-IsExitSave": options.isExitSave ? "true" : "false",
    };
    if (this.#lastModifiedTime) headers["X-COOL-WOPI-Timestamp"] = this.#lastModifiedTime;
    const tried = new Set<string>();
    let tokenUser: UserSession | null = options.tokenUser ?? this.#leader ?? this.#pickToken([]);
    if (tokenUser && !tokenUser.isTokenValidAt(this.#now())) {
      tried.add(tokenUser.id);
      tokenUser = this.#pickToken([], tried);
    }
    while (tokenUser) {
      const current = tokenUser;
      try {
        const result = await this.#host.putFile({
          wopiSrc: this.wopiSrc,
          accessToken: current.accessToken,
          body: bytes,
          lock: this.#lockHeld ? this.lockId : undefined,
          extraHeaders: headers,
        });
        this.#itemVersion = result.itemVersion ?? this.#itemVersion;
        this.#savedVersion = Math.max(this.#savedVersion, versionAtStart);
        if (!this.dirty) this.#firstDirtyAt = null;
        try {
          const info = await this.#host.checkFileInfo({ wopiSrc: this.wopiSrc, accessToken: current.accessToken });
          if (info.LastModifiedTime) this.#lastModifiedTime = info.LastModifiedTime;
          if (info.Version) this.#itemVersion = info.Version;
        } catch {}
        this.#recordWritten();
        const by = options.by && options.by === current.id ? options.by : current.id;
        this.#emit("saved", { itemVersion: this.#itemVersion, by });
        return { saved: true, itemVersion: this.#itemVersion, by };
      } catch (err) {
        const kind = err instanceof WopiHostError ? err.kind : "unknown";
        if (kind === "unauthorized") {
          tried.add(current.id);
          this.#markTokenDead(current);
          tokenUser = this.#pickToken([], tried);
          if (tokenUser) continue;
          this.#emit("saveFailed", { kind });
          this.#log("error", `save failed (${kind})`);
          return { saved: false, itemVersion: this.#itemVersion, by: current.id };
        }
        if (kind === "conflict" || kind === "lockMismatch") {
          return this.#saveConflictCopy(bytes, versionAtStart, current, options, tried);
        }
        this.#emit("saveFailed", { kind });
        this.#log("error", this.#scrub(`save failed (${kind})`));
        return { saved: false, itemVersion: this.#itemVersion, by: current.id };
      }
    }
    this.#emit("saveFailed", { kind: "unauthorized" });
    this.#log("warn", "save failed (unauthorized)");
    return { saved: false, itemVersion: this.#itemVersion, by: null };
  }

  async #saveConflictCopy(bytes: Uint8Array, versionAtStart: number, tokenUser: UserSession, options: SaveOptions, tried: Set<string>): Promise<SaveResult> {
    const target = conflictTargetName(this.#fileName, this.#now());
    let current: UserSession | null = tokenUser;
    while (current) {
      const candidate = current;
      try {
        const result = await this.#host.putRelativeFile({
          wopiSrc: this.wopiSrc,
          accessToken: candidate.accessToken,
          body: bytes,
          suggestedTarget: target,
          lock: this.#lockHeld ? this.lockId : undefined,
        });
        this.#savedVersion = Math.max(this.#savedVersion, versionAtStart);
        if (!this.dirty) this.#firstDirtyAt = null;
        try {
          this.#deps.journal?.flush();
        } catch {}
        this.#emit("conflict", { name: result.name, url: result.url, copyName: result.name, copyUrl: result.url });
        this.#log("warn", "save conflict, wrote a conflict copy", { name: result.name });
        return { saved: true, itemVersion: this.#itemVersion, by: candidate.id, conflictName: result.name, conflictUrl: result.url };
      } catch (err) {
        const kind = err instanceof WopiHostError ? err.kind : "unknown";
        if (kind === "unauthorized") {
          tried.add(candidate.id);
          this.#markTokenDead(candidate);
          current = this.#pickToken([], tried);
          if (current) continue;
        }
        this.#emit("saveFailed", { kind });
        this.#log("error", `save failed (${kind})`);
        return { saved: false, itemVersion: this.#itemVersion, by: candidate.id };
      }
    }
    return { saved: false, itemVersion: this.#itemVersion, by: tokenUser.id };
  }

  async #maybeCloseAfterDeparture(departing: UserSession[]): Promise<void> {
    if (this.#state !== "open") return;
    if (this.#connectedEditors().length > 0) return;
    if (!this.#everHadEditor) {
      const connected = [...this.#users.values()].filter((user) => user.connected);
      if (connected.length === 0) await this.#closeFinal(true);
      return;
    }
    for (;;) {
      if (this.#state !== "open") return;
      if (this.#connectedEditors().length > 0) return;
      if (this.dirty) {
        const tokenUser = this.#pickToken(departing);
        if (!tokenUser) {
          this.#enterHold();
          return;
        }
        const result = await this.#saveInternal({ tokenUser, isAutosave: false, isExitSave: true, by: tokenUser.id });
        if (this.#connectedEditors().length > 0) return;
        if (this.dirty) {
          if (result.saved) continue;
          this.#enterHold();
          return;
        }
        continue;
      }
      if (this.#lockHeld) {
        const unlocked = await this.#unlockWith(departing);
        if (this.#connectedEditors().length > 0) {
          if (unlocked) {
            const editor = this.#connectedEditors()[0];
            if (editor) {
              try {
                await this.#attemptLock(editor.accessToken);
              } catch {
                this.#log("warn", "relock after reconnect failed");
              }
            }
          }
          return;
        }
        if (this.dirty) continue;
      }
      if (this.#connectedEditors().length > 0) return;
      if (this.dirty) continue;
      await this.#closeFinal(true);
      return;
    }
  }

  async #unlockWith(preferred: UserSession[]): Promise<boolean> {
    const tried = new Set<string>();
    let tokenUser = this.#pickToken(preferred);
    while (tokenUser) {
      const current = tokenUser;
      try {
        await this.#host.unlock({ wopiSrc: this.wopiSrc, accessToken: current.accessToken, lock: this.lockId });
        this.#lockHeld = false;
        return true;
      } catch (err) {
        const kind = err instanceof WopiHostError ? err.kind : "unknown";
        this.#log("warn", `unlock failed (${kind})`);
        if (kind !== "unauthorized") return false;
        tried.add(current.id);
        this.#markTokenDead(current);
        tokenUser = this.#pickToken([], tried);
      }
    }
    return false;
  }

  #enterHold(): void {
    this.#holdPending = true;
    const base = this.#lastLockRefreshAt || this.#now();
    const delay = Math.max(0, base + HOLD_FLUSH_MS - this.#now());
    this.#clearTimer("hold");
    this.#holdTimer = this.#scheduler.setTimeout(() => {
      this.#enqueueTask(async () => {
        if (!this.#holdPending || this.#state !== "open") return;
        if (this.#connectedEditors().length > 0) return;
        this.#holdPending = false;
        this.#log("error", "lost flush after hold window");
        this.#emit("lostFlush", {});
        await this.#closeFinal(false);
      });
    }, delay);
    this.#log("warn", "no valid token for final flush, holding the session");
  }

  async #closeFinal(deleteJournal: boolean): Promise<void> {
    if (this.#state === "closed") return;
    this.#state = "closing";
    this.#clearTimer("refresh");
    this.#clearTimer("autosave");
    this.#clearTimer("hold");
    if (this.#sweepTimer !== null) {
      this.#scheduler.clearInterval(this.#sweepTimer);
      this.#sweepTimer = null;
    }
    for (const id of [...this.#userTimers.keys()]) this.#clearUserTimers(id);
    if (deleteJournal && !this.dirty && this.#journalActive) {
      try {
        this.#deps.journal?.delete();
      } catch {}
    } else {
      this.#deps.journal?.close();
    }
    this.#destroyReplica();
    for (const user of this.#history) user.accessToken = "";
    this.#users.clear();
    this.#history = [];
    this.#leader = null;
    this.#lockHeld = false;
    this.#state = "closed";
    this.#emit("closed", {});
    this.#deps.onClosed?.(this);
  }

  /** Drops timers and the replica without any host call. Used to simulate a crash. */
  abandon(): void {
    if (this.#state === "closed") return;
    this.#clearAllTimers();
    this.#deps.journal?.close();
    this.#destroyReplica();
    for (const user of this.#history) user.accessToken = "";
    this.#users.clear();
    this.#state = "closed";
    this.#deps.onClosed?.(this);
  }

  /** Closes the session without flushing. Prefer leave for normal shutdown. */
  close(): Promise<void> {
    return this.#enqueue(() => this.#closeFinal(false));
  }
}

