import type { SlideDisplayList, SlidePrimitive } from '../types';

export function frameAssetIds(displayList: SlideDisplayList): Set<string> {
  const ids = new Set<string>();
  const visit = (primitives: readonly SlidePrimitive[]): void => {
    for (const primitive of primitives) {
      if (primitive.kind === 'image' && primitive.assetId) ids.add(primitive.assetId);
      else if (primitive.kind === 'chart' || primitive.kind === 'table') visit(primitive.primitives);
    }
  };
  visit(displayList.primitives);
  return ids;
}
