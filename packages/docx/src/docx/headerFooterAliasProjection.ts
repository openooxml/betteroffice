import type { HeaderFooter } from '../types/document';

const aliases = new WeakMap<HeaderFooter, string>();

/** @internal */
export function markHeaderFooterAlias(entry: HeaderFooter, canonicalRId: string): void {
  aliases.set(entry, canonicalRId);
}

/** @internal */
export function headerFooterAliasOf(entry: HeaderFooter): string | undefined {
  return aliases.get(entry);
}

/** @internal */
export function isWrittenByCanonical(
  entry: HeaderFooter,
  entries: ReadonlyMap<string, HeaderFooter>
): boolean {
  const canonicalRId = headerFooterAliasOf(entry);
  return canonicalRId !== undefined && entries.has(canonicalRId);
}
