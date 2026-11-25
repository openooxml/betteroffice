import type { FontPackDiagnostic } from './fontPackDiagnostics';

export type FontCdnDiagnostic = Omit<FontPackDiagnostic, 'reason'> & { reason: 'cdn-unavailable' };

export function isImplicitFontCdnError(error: unknown): error is Error & { diagnostic: FontCdnDiagnostic } {
  return error instanceof Error &&
    (error as { diagnostic?: FontCdnDiagnostic }).diagnostic?.reason === 'cdn-unavailable';
}

export function reportImplicitCdnFailure({ diagnostic }: { diagnostic: FontCdnDiagnostic }): void {
  if (typeof globalThis.dispatchEvent === 'function' && typeof CustomEvent === 'function') {
    globalThis.dispatchEvent(new CustomEvent('betteroffice:font-pack-diagnostic', { detail: { ...diagnostic } }));
  }
}
