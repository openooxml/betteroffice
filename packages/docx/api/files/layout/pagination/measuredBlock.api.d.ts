// @betteroffice/docx declaration layout/pagination/measuredBlock.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { BlockExtent, LayoutBlock } from './types';
export interface MeasuredBlock {
    block: LayoutBlock;
    measure: BlockExtent;
}
