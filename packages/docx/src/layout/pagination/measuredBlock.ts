/**
 * The colocated block+measure unit consumed by the paginator.
 */

import type { LayoutBlock, BlockExtent } from './types';

/**
 * A flow block paired with its measurement. The paginator consumes
 * `MeasuredBlock[]` instead of two parallel `LayoutBlock[]` + `BlockExtent[]`
 * arrays kept in lockstep by index — colocating the pair makes a misaligned
 * block/measure unrepresentable.
 * @public
 */
export type MeasuredBlock = {
  block: LayoutBlock;
  measure: BlockExtent;
};

/**
 * Zip aligned `blocks` and `measures` into `MeasuredBlock[]`. This is the one
 * place a length assumption is localized — the pervasive per-index coupling it
 * replaces is what the measured-block model removes.
 * @public
 */
export function toMeasuredBlocks(blocks: LayoutBlock[], measures: BlockExtent[]): MeasuredBlock[] {
  if (blocks.length !== measures.length) {
    throw new Error(
      `toMeasuredBlocks: expected one measure per block (blocks=${blocks.length}, measures=${measures.length})`
    );
  }
  return blocks.map((block, i) => ({ block, measure: measures[i] }));
}
