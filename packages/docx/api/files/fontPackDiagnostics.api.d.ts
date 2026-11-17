// @betteroffice/docx declaration fontPackDiagnostics.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { FontPackDiagnostic } from '../../../shared/fontPackDiagnostics';
export type { FontPackDiagnostic };
/** The installed font add-on packages that cannot be used, and why; their faces are substituted. */
export declare function fontPackDiagnostics(): Promise<FontPackDiagnostic[]>;
