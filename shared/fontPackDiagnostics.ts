/** A font pack that cannot be used, so its faces are substituted. */
export interface FontPackDiagnostic {
  /** The add-on package: `@betteroffice/fonts-latin`, `-intl` or `-cjk`. */
  pack: 'latin' | 'intl' | 'cjk';
  reason: 'version-mismatch' | 'missing-file' | 'invalid-manifest' | 'cdn-unavailable';
  /** A sentence naming the package and what is wrong. */
  detail: string;
}

export function isImplicitFontCdnError(error: unknown): error is Error & { diagnostic: FontPackDiagnostic } {
  return error instanceof Error &&
    (error as { diagnostic?: FontPackDiagnostic }).diagnostic?.reason === 'cdn-unavailable';
}

const CDN_WARNED = Symbol.for('betteroffice.fonts.cdn-warned');

export function reportImplicitCdnFailure({ diagnostic }: { diagnostic: FontPackDiagnostic }): void {
  const scope = globalThis as { [CDN_WARNED]?: Set<FontPackDiagnostic['pack']> };
  const warned = scope[CDN_WARNED] ??= new Set();
  if (!warned.has(diagnostic.pack)) {
    warned.add(diagnostic.pack);
    console.warn(diagnostic.detail);
  }
  if (typeof globalThis.dispatchEvent === 'function' && typeof CustomEvent === 'function') {
    globalThis.dispatchEvent(new CustomEvent('betteroffice:font-pack-diagnostic', { detail: { ...diagnostic } }));
  }
}

const WARNED = Symbol.for('@betteroffice/font-pack-warnings');

/**
 * Emits each diagnostic once per pack and process as a Node process warning with code `BETTEROFFICE_FONT_PACK`,
 * so server hosts can filter it. Does nothing outside Node.
 */
export function warnFontPackDiagnostics(diagnostics: readonly FontPackDiagnostic[]): void {
  const proc = (globalThis as {
    process?: { emitWarning?: (warning: string, options: { type: string; code: string }) => void };
  }).process;
  if (typeof proc?.emitWarning !== 'function') return;
  const store = globalThis as { [WARNED]?: Set<string> };
  const warned = (store[WARNED] ??= new Set());
  for (const { pack, detail } of diagnostics) {
    if (warned.has(pack)) continue;
    warned.add(pack);
    proc.emitWarning(detail, { type: 'BetterOfficeFontPackWarning', code: 'BETTEROFFICE_FONT_PACK' });
  }
}
