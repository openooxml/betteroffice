const HAS_SCHEME = /^[a-z][a-z0-9+.-]*:/i;
const ALLOWED_SCHEME = /^(https?|mailto|tel|ftp):/i;

/**
 * DOCX keeps relative hrefs. PPTX slide links have no base document and use `relative: false`.
 * XLSX also drops relative hrefs, allows only HTTP, HTTPS, mailto and tel, and normalizes URLs at its sink.
 */
export function sanitizeHref(
  href: string | null | undefined,
  options: { relative?: boolean; protocols?: readonly string[] } = {},
): string | undefined {
  if (!href) return undefined;
  const probe = href.replace(/[\t\n\r]/g, '').replace(/^[\x00-\x20]+/, '');
  if (!probe) return undefined;
  const scheme = HAS_SCHEME.exec(probe)?.[0].toLowerCase();
  if (!scheme) return options.relative === false ? undefined : href;
  if (options.protocols ? !options.protocols.includes(scheme) : !ALLOWED_SCHEME.test(probe)) return undefined;
  if (options.relative === false) {
    try { new URL(probe); } catch { return undefined; }
  }
  return href;
}
