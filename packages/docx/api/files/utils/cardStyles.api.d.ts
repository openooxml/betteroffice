// @betteroffice/docx declaration utils/cardStyles.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Sidebar card chrome — shared between React and Vue. Numeric
 * pixel values rather than `'8px'` strings so both adapters'
 * CSSProperties shapes accept them. Lifted from
 * `packages/docx-react/src/components/sidebar/cardStyles.ts` and the
 * Vue mirror so there's one canonical table.
 * @packageDocumentation
 * @public
 */
import type { CSSProperties } from './cssTypes';
export declare const CARD_STYLE_COLLAPSED: CSSProperties;
export declare const CARD_STYLE_EXPANDED: CSSProperties;
