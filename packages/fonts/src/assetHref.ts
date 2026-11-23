export function absoluteHref(href: string): string {
  try {
    new URL(href);
    return href;
  } catch {
    try {
      return new URL(href, globalThis.location?.href).href;
    } catch {
      return href;
    }
  }
}
