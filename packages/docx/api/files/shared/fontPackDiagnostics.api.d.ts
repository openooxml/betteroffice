// @betteroffice/docx declaration shared/fontPackDiagnostics.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** An installed font pack that cannot be used, so its faces are substituted. */
export interface FontPackDiagnostic {
    /** The add-on package: `@betteroffice/fonts-latin`, `-intl` or `-cjk`. */
    pack: 'latin' | 'intl' | 'cjk';
    /** `version-mismatch`: another release is installed; `missing-file`: files are absent; `invalid-manifest`: its manifest cannot be read. */
    reason: 'version-mismatch' | 'missing-file' | 'invalid-manifest';
    /** A sentence naming the package and what is wrong. */
    detail: string;
}
/**
 * Emits each diagnostic once per pack and process as a Node process warning with code `BETTEROFFICE_FONT_PACK`,
 * so server hosts can filter it. Does nothing outside Node.
 */
export declare function warnFontPackDiagnostics(diagnostics: readonly FontPackDiagnostic[]): void;
