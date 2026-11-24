import type { FontPackDiagnostic } from './fontPackDiagnostics';

export type FontCdnDiagnostic = Omit<FontPackDiagnostic, 'reason'> & { reason: 'cdn-unavailable' };

const REPORTED = Symbol.for('betteroffice.fonts.cdn.reported-optional-failures');

export function isImplicitFontCdnError(error: unknown): error is Error & { diagnostic: FontCdnDiagnostic } {
  return error instanceof Error &&
    (error as { diagnostic?: FontCdnDiagnostic }).diagnostic?.reason === 'cdn-unavailable';
}

export function reportImplicitCdnFailure(error: { diagnostic: FontCdnDiagnostic }): void {
  if (typeof globalThis.dispatchEvent !== 'function' || typeof CustomEvent !== 'function') return;
  const { diagnostic } = error;
  if (diagnostic.pack === 'cjk' || diagnostic.pack === 'intl') {
    const scope = globalThis as { [REPORTED]?: WeakSet<object> };
    const reported = (scope[REPORTED] ??= new WeakSet());
    if (reported.has(error)) return;
    reported.add(error);
  }
  globalThis.dispatchEvent(new CustomEvent('betteroffice:font-pack-diagnostic', { detail: { ...diagnostic } }));
}
