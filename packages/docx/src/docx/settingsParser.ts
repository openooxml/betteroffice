import type { EndnoteProperties, FootnoteProperties } from '../types/document';

export interface CompatibilityFlags {
  compatibilityMode: number;
  noLeading: boolean;
  doNotExpandShiftReturn: boolean;
  useWord97LineBreakRules: boolean;
  balanceSingleByteDoubleByteWidth: boolean;
}

export const DEFAULT_COMPATIBILITY_FLAGS: CompatibilityFlags = {
  compatibilityMode: 12,
  noLeading: false,
  doNotExpandShiftReturn: false,
  useWord97LineBreakRules: false,
  balanceSingleByteDoubleByteWidth: false,
};

/** Public model contract rehydrated by the Rust S9 facade. */
export interface DocumentSettings {
  defaultTabStop: number;
  defaultTableStyle?: string;
  themeFontLang?: { eastAsia?: string; bidi?: string };
  compatibilityFlags: CompatibilityFlags;
  updateFields?: boolean;
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

export const DEFAULT_TAB_STOP_TWIPS = 720;
