import type { FontPackDiagnostic } from '../../../shared/fontPackDiagnostics';

export type { FontPackDiagnostic };

/** The installed font add-on packages that cannot be used, and why; their faces are substituted. */
export async function fontPackDiagnostics(): Promise<FontPackDiagnostic[]> {
  return (await import('@betteroffice/fonts')).checkFontPacks();
}
