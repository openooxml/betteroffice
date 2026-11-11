// @betteroffice/docx declaration layout/measure/rustMeasureSource.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { ProgressiveFontSpec } from '../../yrs/residentProgressiveFonts';
import type { CompatibilityFlags } from '../../docx/settingsParser';
import type { FontDecoder } from '../../../../../shared/office-session/fontDecoder';
import { type BundledFontProvider, type EmbeddedFaceInput, type FontScript } from './fontRegistry';
export interface RustTextEngine {
    registerFont(bytes: Uint8Array, face?: {
        family: string;
        bold: boolean;
        italic: boolean;
    }): number;
    /**
     * Measurement view of `id` carrying the vertical metrics and advance pitch
     * Word measures `family` with; `id` when the engine knows none. Optional so
     * a host that predates it still measures, with the substitute's own
     * metrics.
     */
    registerSubstituteFont?(id: number, family: string): number;
    clearFonts(): void;
}
export interface ResidentFontRequirement {
    key: string;
    family: string;
    bold: boolean;
    italic: boolean;
    scripts?: FontScript[];
    characters?: string[];
    characterBits?: string;
}
/** @internal */
export declare function fontRequirementCharacters(requirement: ResidentFontRequirement): string[];
export interface ResidentMeasurementConfig {
    fontChains: Record<string, number[]>;
    defaults: {
        fontSize: number;
        fontFamily: string;
    };
    compat: {
        compatibilityMode?: number;
        noLeading: boolean;
        doNotExpandShiftReturn: boolean;
    };
    authoritativeShaping: true;
}
export interface RustMeasureSource {
    setEmbeddedFaces(faces: EmbeddedFaceInput[], appFaces?: EmbeddedFaceInput[]): void;
    setCompat(flags: CompatibilityFlags | undefined): void;
    prepareFontRequirements(requirements: ResidentFontRequirement[]): Promise<boolean>;
    measurementConfigForRequirements(requirements: ResidentFontRequirement[]): ResidentMeasurementConfig | undefined;
    /** @internal Whether later requirements need no face the prepared chains lack (characters the base Latin tier settled count as covered); holds no font bytes, so it outlives `clear()`. */
    coverageCheck(requirements: ResidentFontRequirement[]): (next: ResidentFontRequirement[]) => boolean;
    clear(): void;
    /** The script-fallback slots registered from a sliced face's skeleton and core, with their files. */
    progressiveFonts?(): Promise<ProgressiveFontSpec[]>;
}
export declare function getRustTextEngine(): Promise<RustTextEngine>;
/** Builds a source with an explicit or default font provider. */
export declare function createRustMeasureSource(options: {
    engine: RustTextEngine;
    bundled?: BundledFontProvider;
    runtime?: {
        decoder?: FontDecoder;
    };
    /** Register sliced script-fallback faces from their skeleton and core. */
    progressive?: boolean;
    /** With the default provider, give families it doesn't carry the last-resort face. */
    standIns?: boolean;
}): RustMeasureSource;
