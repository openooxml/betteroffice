// @betteroffice/docx declaration shared/office-session/pngExportError.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
export type PngExportCode = 'canvas-unavailable' | 'canvas-too-large' | 'font-unavailable' | 'picture-unavailable';
export declare class PngExportError extends Error {
    readonly code: PngExportCode;
    constructor(code: PngExportCode, message: string);
}
/** Why a picture was left as a placeholder. */
export type PngPictureUnavailableReason = 'missing' | 'unsupported' | 'decode-failed' | 'metafile' | 'too-large' | 'budget';
/** A picture a PNG export drew as a placeholder. `limit` and `actual` say by how much a size or budget limit was passed. */
export interface PngUnavailablePicture {
    /** The engine's reference for the picture; empty for one whose relationship has no part, which has nothing to name it by. */
    id: string;
    reason: PngPictureUnavailableReason;
    limit?: number;
    actual?: number;
    unit?: 'pixels' | 'bytes' | 'work';
}
/**
 * The kinds of effect a PNG export can leave off a page, in the order an
 * `effects-skipped` diagnostic lists them:
 * - `text.outline`: a run's classic hollow outline.
 * - `text.emphasisMark`: the dots, commas or circles marking emphasised characters.
 * - `text.effect`: a legacy animated text effect such as a blinking background.
 * - `text.modernEffects`: glow, reflection, gradient fill, outline and 3-D text effects.
 * - `text.shadow`: a run's shadow, emboss or engrave the engine does not draw.
 * - `shape.effects`: a shape effect past the one shadow or glow the engine draws.
 */
export declare const PNG_SKIPPED_EFFECT_KINDS: readonly ["text.outline", "text.emphasisMark", "text.effect", "text.modernEffects", "text.shadow", "shape.effects"];
export type PngSkippedEffectKind = (typeof PNG_SKIPPED_EFFECT_KINDS)[number];
/** One kind of effect a page was drawn without, and how many times. */
export interface PngSkippedEffect {
    /** One of {@link PNG_SKIPPED_EFFECT_KINDS}; typed as a string so a later kind is not a breaking change. */
    kind: string;
    count: number;
}
/** Pictures a PNG export drew as a placeholder. */
export interface PngPictureUnavailableDiagnostic {
    code: 'picture-unavailable';
    pictures: PngUnavailablePicture[];
}
/**
 * Effects a PNG export left off a page, sent at most once per page and only
 * when it skipped something. `page` is zero-based, as the export target is.
 * The counts are what the engine skipped, not a statement about pixels:
 * counts are drawing primitives on the page that lost the effect, not runs
 * or objects. A picture that is not drawn reports `picture-unavailable` only.
 */
export interface PngEffectsSkippedDiagnostic {
    code: 'effects-skipped';
    page: number;
    effects: PngSkippedEffect[];
}
/** A problem a PNG export worked around: the page still exports. */
export type PngDiagnostic = PngPictureUnavailableDiagnostic | PngEffectsSkippedDiagnostic;
/** Orders counted kinds as {@link PNG_SKIPPED_EFFECT_KINDS} does, dropping zero counts; throws on a kind the list lacks. */
export declare function skippedEffects(counts: ReadonlyMap<string, number>): PngSkippedEffect[];
