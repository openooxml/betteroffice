/**
 * Which table style a freshly inserted table should wear.
 *
 * Both insert paths (toolbar and agent) call into here so a new table picks
 * up the document's own table look instead of a bare black grid. The
 * document can express that look two ways, and we trust them in order of
 * specificity: an explicit `w:defaultTableStyle` in settings.xml
 * (§17.15.1.44) wins when it points at a real style; failing that, a table
 * style marked `w:default="1"` (§17.7.4.18) is used — except when that
 * default is only the formatting-free base ("Table Normal"), which every
 * table already inherits and which would add nothing. With neither present
 * the caller keeps its plain-border fallback.
 *
 * We deliberately never guess by probing for the TableGrid builtin: a
 * document may ship that style without meaning it as its insert default,
 * and the two spec elements above are how templates state intent.
 */

import type { TableLook } from '../types/document';
import type { StyleResolver } from './styleResolver';

/**
 * `tblLook` Word applies to a brand-new table (val="04A0"): style the first
 * row and first column, enable horizontal banding, disable vertical banding.
 * Setting this lets a resolved table style paint its header-row and banding
 * conditional formatting on the inserted table, matching how Word renders a
 * fresh table created with that style.
 */
export const DEFAULT_NEW_TABLE_LOOK: TableLook = {
  firstRow: true,
  firstColumn: true,
  lastRow: false,
  lastColumn: false,
  noHBand: false,
  noVBand: true,
};

/** Whether a style is the no-op base table style every table inherits from. */
function isBaseTableStyle(styleId: string, name: string | undefined): boolean {
  if (styleId.toLowerCase() === 'tablenormal') return true;
  const normalized = name?.trim().toLowerCase();
  return normalized === 'table normal' || normalized === 'normal table';
}

/**
 * StyleId a brand-new table should adopt, or `undefined` when the document
 * declares nothing usable (the caller then keeps its plain-border default).
 *
 * @param settingsStyleId - `w:defaultTableStyle` from settings.xml
 * @param styles - the document's style resolver (null when unavailable)
 */
export function pickTableStyleForInsert(
  settingsStyleId: string | undefined | null,
  styles: StyleResolver | null | undefined
): string | undefined {
  if (!styles) return undefined;

  // a settings.xml declaration is authoritative — but only when it points at
  // a style that exists; serializing a dangling <w:tblStyle> would be invalid
  if (settingsStyleId && styles.hasStyle(settingsStyleId)) {
    return settingsStyleId;
  }

  const typeDefault = styles.getDefaultTableStyle();
  if (!typeDefault?.styleId) return undefined;

  // the w:default="1" style counts only when it actually formats something —
  // the bare base style is what every table inherits anyway
  return isBaseTableStyle(typeDefault.styleId, typeDefault.name) ? undefined : typeDefault.styleId;
}
