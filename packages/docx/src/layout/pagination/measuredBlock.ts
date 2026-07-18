import type { BlockExtent, LayoutBlock } from './types';

export interface MeasuredBlock {
  block: LayoutBlock;
  measure: BlockExtent;
}
