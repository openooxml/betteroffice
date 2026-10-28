import { randomBytes } from "node:crypto";

/** Inputs used to create one user handle. */
export interface UserSessionInit {
  hostId: string;
  wopiSrc: string;
  accessToken: string;
  accessTokenTtl: number;
  userId: string;
  friendlyName: string;
  canWrite: boolean;
  joinedAt: number;
  postMessageOrigin?: string;
  closeUrl?: string;
  closePostMessage?: boolean;
}

/** Per user handle inside a file session. The token is never serialized. */
export class UserSession {
  readonly id: string;
  readonly hostId: string;
  readonly wopiSrc: string;
  accessToken: string;
  expiresAt: number | null;
  readonly userId: string;
  readonly friendlyName: string;
  canWrite: boolean;
  readonly joinedAt: number;
  readonly postMessageOrigin: string;
  readonly closeUrl: string;
  readonly closePostMessage: boolean;
  lastSeen: number;
  connected: boolean;

  constructor(init: UserSessionInit) {
    this.id = randomBytes(32).toString("base64url");
    this.hostId = init.hostId;
    this.wopiSrc = init.wopiSrc;
    this.accessToken = init.accessToken;
    this.expiresAt = init.accessTokenTtl === 0 ? null : init.accessTokenTtl;
    this.userId = init.userId;
    this.friendlyName = init.friendlyName;
    this.canWrite = init.canWrite;
    this.joinedAt = init.joinedAt;
    this.postMessageOrigin = init.postMessageOrigin ?? "";
    this.closeUrl = init.closeUrl ?? "";
    this.closePostMessage = init.closePostMessage === true;
    this.lastSeen = init.joinedAt;
    this.connected = true;
  }

  /** Whether the token is known to still be valid at the given time. */
  isTokenValidAt(now: number): boolean {
    return this.expiresAt === null || this.expiresAt > now;
  }

  toJSON(): Record<string, unknown> {
    return {
      id: this.id,
      hostId: this.hostId,
      wopiSrc: this.wopiSrc,
      expiresAt: this.expiresAt,
      userId: this.userId,
      friendlyName: this.friendlyName,
      canWrite: this.canWrite,
      joinedAt: this.joinedAt,
      postMessageOrigin: this.postMessageOrigin,
      closeUrl: this.closeUrl,
      closePostMessage: this.closePostMessage,
      lastSeen: this.lastSeen,
      connected: this.connected,
    };
  }

  toString(): string {
    return `UserSession(${this.id})`;
  }

  [Symbol.for("nodejs.util.inspect.custom")](): string {
    return this.toString();
  }
}
