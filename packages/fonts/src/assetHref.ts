export function absoluteHref(href: string): string {
  try {
    return new URL(href).href;
  } catch {
    try {
      return new URL(href, globalThis.location?.href).href;
    } catch {
      return href;
    }
  }
}
