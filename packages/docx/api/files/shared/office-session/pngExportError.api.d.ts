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
 * The kinds of effect a PNG export can leave off a page:
 * - `text.outline`: a run's classic hollow outline.
 * - `text.emphasisMark`: the dots, commas or circles marking emphasised characters.
 * - `text.effect`: a legacy animated text effect such as a blinking background.
 * - `text.modernEffects`: a run's glow, shadow, reflection, fill or outline text effects.
 * - `text.shadow`: a run's classic shadow, emboss or engrave.
 * - `text.leaderGlyphs`: the dots or dashes a tab leader repeats.
 * - `shape.effects`: a shape's effects, such as a soft edge or a second shadow.
 * - `image.filter`: a picture's colour filter, such as a watermark washout.
 * - `image.effects`: a picture's colour or blur effects.
 *
 * A page reports a kind when it carries at least one drawing primitive with an
 * effect of that kind that the engine does not draw on its normal path. A
 * picture whose bytes cannot be loaded or decoded reports `picture-unavailable`
 * and nothing else. Kinds come in the order listed here.
 *
 * The report describes each engine's normal path, not the pixels of one page.
 * A degenerate primitive that carries an effect, such as an empty crop, a run
 * of spaces or a leader with no advance, still reports it.
 */
export declare const PNG_SKIPPED_EFFECT_KINDS: readonly ["text.outline", "text.emphasisMark", "text.effect", "text.modernEffects", "text.shadow", "text.leaderGlyphs", "shape.effects", "image.filter", "image.effects"];
export type PngSkippedEffectKind = (typeof PNG_SKIPPED_EFFECT_KINDS)[number];
/** Pictures a PNG export drew as a placeholder. */
export interface PngPictureUnavailableDiagnostic {
    code: 'picture-unavailable';
    pictures: PngUnavailablePicture[];
}
/**
 * Effects a PNG export left off a page, sent at most once per page and only
 * when `effects` is not empty. `page` is zero-based, as the export target is.
 * `effects` names each kind once, in {@link PNG_SKIPPED_EFFECT_KINDS} order.
 */
export interface PngEffectsSkippedDiagnostic {
    code: 'effects-skipped';
    page: number;
    effects: PngSkippedEffectKind[];
}
/** A problem a PNG export worked around: the page still exports. */
export type PngDiagnostic = PngPictureUnavailableDiagnostic | PngEffectsSkippedDiagnostic;
