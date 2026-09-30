/** A mouse press on a page surface needs the replica; plugin overlays and touch or pen pans do not. */
export function pagePressNeedsReplica(event: PointerEvent): boolean {
  if (event.pointerType === 'touch' || event.pointerType === 'pen') return false;
  return event.target instanceof Element && event.target.closest('.canvas-page') !== null;
}
