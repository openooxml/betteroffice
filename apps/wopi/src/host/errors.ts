/** Kinds of failure a WOPI host call can report. */
export type WopiHostErrorKind =
  | "unauthorized"
  | "notFound"
  | "lockMismatch"
  | "conflict"
  | "tooLarge"
  | "preconditionFailed"
  | "notImplemented"
  | "redirect"
  | "timeout"
  | "network"
  | "server"
  | "notAllowed"
  | "invalidResponse";

export interface WopiHostErrorOptions {
  status?: number;
  currentLock?: string | null;
  reason?: string;
  cause?: unknown;
}

/** Typed WOPI host failure. Messages never carry an access token. */
export class WopiHostError extends Error {
  readonly kind: WopiHostErrorKind;
  readonly status?: number;
  readonly currentLock?: string | null;
  readonly reason?: string;

  constructor(kind: WopiHostErrorKind, message: string, options: WopiHostErrorOptions = {}) {
    super(message, options.cause !== undefined ? { cause: options.cause } : undefined);
    this.name = "WopiHostError";
    this.kind = kind;
    if (options.status !== undefined) this.status = options.status;
    if (options.currentLock !== undefined) this.currentLock = options.currentLock;
    if (options.reason !== undefined) this.reason = options.reason;
  }
}
