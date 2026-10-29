import { WopiHostError } from "./errors";

/** Redacts access_token values from text. */
export function redactText(text: string): string {
  return text
    .replace(/([?&]access_token=)[^&#\s'"]+/gi, "$1***")
    .replace(/(access_token%3D)[^&#\s'"]+/gi, "$1***")
    .replace(/(access_token=)[^&#\s'"]+/gi, "$1***");
}

/** Returns a URL string with the access_token query value replaced. */
export function redactUrl(url: string | URL): string {
  const raw = typeof url === "string" ? url : url.toString();
  try {
    const parsed = new URL(raw);
    for (const key of [...parsed.searchParams.keys()]) {
      if (key.toLowerCase() === "access_token") parsed.searchParams.set(key, "***");
    }
    if (parsed.hash) parsed.hash = redactText(parsed.hash);
    return redactText(parsed.toString());
  } catch {
    return redactText(raw);
  }
}

/** Returns an error whose message cannot contain an access token. */
export function redactError(err: unknown): Error {
  if (err instanceof WopiHostError) {
    return new WopiHostError(err.kind, redactText(err.message), {
      status: err.status,
      currentLock: err.currentLock ? redactText(err.currentLock) : err.currentLock,
      reason: err.reason !== undefined ? redactText(err.reason) : err.reason,
    });
  }
  if (err instanceof Error) {
    const redacted = new Error(redactText(err.message));
    redacted.name = err.name;
    return redacted;
  }
  if (typeof err === "string") return new Error(redactText(err));
  return new Error(redactText(String(err)));
}
