// @betteroffice/docx declaration docx/settingsParser.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { EndnoteProperties, FootnoteProperties } from '../types/document';
export interface CompatibilityFlags {
    compatibilityMode: number;
    noLeading: boolean;
    doNotExpandShiftReturn: boolean;
    useWord97LineBreakRules: boolean;
    balanceSingleByteDoubleByteWidth: boolean;
    doNotUseHTMLParagraphAutoSpacing?: boolean;
    suppressSpBfAfterPgBrk?: boolean;
    allowSpaceOfSameStyleInTable?: boolean;
    adjustLineHeightInTable?: boolean;
}
export declare const DEFAULT_COMPATIBILITY_FLAGS: CompatibilityFlags;
/** Public model contract rehydrated by the Rust parser. */
export interface DocumentSettings {
    defaultTabStop: number;
    defaultTableStyle?: string;
    themeFontLang?: {
        eastAsia?: string;
        bidi?: string;
    };
    compatibilityFlags: CompatibilityFlags;
    updateFields?: boolean;
    evenAndOddHeaders?: boolean;
    trackRevisions?: boolean;
    doNotTrackMoves?: boolean;
    doNotTrackFormatting?: boolean;
    revisionView?: {
        markup?: boolean;
        comments?: boolean;
        insertionsDeletions?: boolean;
        formatting?: boolean;
    };
    footnotePr?: FootnoteProperties;
    endnotePr?: EndnoteProperties;
}
export declare const DEFAULT_TAB_STOP_TWIPS = 720;
