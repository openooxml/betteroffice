import { isFloatingWrapType, isWrapNone } from '../../docx/wrapTypes';
import type { TextBoxBlock } from './types';

/**
 * Subset of {@link TextBoxBlock} needed to classify how a text box flows
 * relative to surrounding content. Kept narrow so callers (measure, layout,
 * paint) can pass partial views without rebuilding the full block.
 *
 * @public
 */
export type TextBoxFlowAttrs = Pick<TextBoxBlock, 'displayMode' | 'wrapType'>;

/**
 * True when a text box participates in float layout — either via the
 * `float` display mode (CSS-style) or via an OOXML `wrapType` that
 * positions the box outside paragraph flow (`square`, `tight`, `through`,
 * `behind`, `inFront`, `topAndBottom`).
 *
 * `topAndBottom` is positioned (it honours its OOXML anchor, e.g. a banner
 * pinned to the top of the page) even though text never wraps beside it —
 * see {@link floatingTextBoxReservesBand}.
 *
 * @public
 */
export function isFloatingTextBoxBlock(block: TextBoxFlowAttrs): boolean {
  return (
    block.displayMode === 'float' ||
    isFloatingWrapType(block.wrapType) ||
    block.wrapType === 'topAndBottom'
  );
}

/**
 * True when a floating text box reserves a full-width vertical band rather
 * than a side exclusion. `topAndBottom` boxes break the text above and below
 * the box (no text beside it), so surrounding lines flow past the band.
 *
 * @public
 */
export function floatingTextBoxReservesBand(block: TextBoxFlowAttrs): boolean {
  return isFloatingTextBoxBlock(block) && block.wrapType === 'topAndBottom';
}

/**
 * True when a floating text box reserves an exclusion zone that narrows
 * surrounding text lines. `wrapNone` (`behind` / `inFront`) and
 * `topAndBottom` are floats that don't wrap text on their sides.
 *
 * @public
 */
export function floatingTextBoxWrapsText(block: TextBoxFlowAttrs): boolean {
  return (
    isFloatingTextBoxBlock(block) &&
    !isWrapNone(block.wrapType) &&
    block.wrapType !== 'topAndBottom'
  );
}
