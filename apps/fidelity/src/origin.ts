/** Resolves a query override against the page, keeping it only as an absolute URL on the same origin. */
export function sameOrigin(value: string | null, base: string): string | null {
  if (!value) return null;
  try {
    const url = new URL(value, base);
    return url.origin === new URL(base).origin ? url.href : null;
  } catch {
    return null;
  }
}
