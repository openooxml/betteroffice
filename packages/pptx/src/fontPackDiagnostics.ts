import type { FontPackDiagnostic } from '../../../shared/fontPackDiagnostics';
import { fontPackDiagnostics as installedPackDiagnostics } from './wasm/engine';

export type { FontPackDiagnostic };

/** The installed font add-on packages that cannot be used, and why; their faces are substituted. */
export function fontPackDiagnostics(): Promise<FontPackDiagnostic[]> {
  return installedPackDiagnostics();
}
