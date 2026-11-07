/** An installed font pack that cannot be used, so its faces are substituted. */
export interface FontPackDiagnostic {
  /** The add-on package: `@betteroffice/fonts-latin`, `-intl` or `-cjk`. */
  pack: 'latin' | 'intl' | 'cjk';
  /** `version-mismatch`: another release is installed; `missing-file`: files are absent; `invalid-manifest`: its manifest cannot be read. */
  reason: 'version-mismatch' | 'missing-file' | 'invalid-manifest';
  /** A sentence naming the package and what is wrong. */
  detail: string;
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
