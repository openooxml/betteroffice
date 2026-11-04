import type { PeerlessFrameGeometry, PeerlessViewSnapshot } from './viewStore';

/** @internal Value equality for shell state and worker query results. */
export function sameShellValue(first: unknown, second: unknown): boolean {
  if (Object.is(first, second)) return true;
  if (first == null || second == null || typeof first !== 'object' || typeof second !== 'object') return false;
  if (first instanceof Uint8Array || second instanceof Uint8Array) {
    return first instanceof Uint8Array && second instanceof Uint8Array && first.length === second.length &&
      first.every((value, index) => value === second[index]);
  }
  if (Array.isArray(first) || Array.isArray(second)) {
    return Array.isArray(first) && Array.isArray(second) && first.length === second.length &&
      first.every((value, index) => sameShellValue(value, second[index]));
  }
  const left = first as Record<string, unknown>;
  const right = second as Record<string, unknown>;
  const keys = Object.keys(left);
  return keys.length === Object.keys(right).length && keys.every(key =>
    Object.prototype.hasOwnProperty.call(right, key) && sameShellValue(left[key], right[key]));
}

function sameGeometry(first: PeerlessFrameGeometry | null, second: PeerlessFrameGeometry | null): boolean {
  if (first === second) return true;
  if (!first || !second) return false;
  if (first.documentVersion !== second.documentVersion || !sameShellValue(first.identity, second.identity) ||
    first.coversThroughSeq !== second.coversThroughSeq || first.layoutRevision !== second.layoutRevision ||
    first.layoutEpoch !== second.layoutEpoch || first.documentWidth !== second.documentWidth ||
    first.layoutProvisional !== second.layoutProvisional || !sameShellValue(first.reservations, second.reservations) ||
    !sameShellValue(first.caret, second.caret) || !sameShellValue(first.selectionFocus, second.selectionFocus) ||
    !sameShellValue(first.selectionRects, second.selectionRects) || !sameShellValue(first.parts, second.parts)) return false;
  const left = first.objects;
  const right = second.objects;
  if (left !== right && (!left || !right || left.version !== right.version || left.layoutRevision !== right.layoutRevision ||
    !sameShellValue(left.selection, right.selection) || !sameShellValue(left.objects, right.objects) ||
    !sameShellValue(left.unbuiltPages, right.unbuiltPages))) return false;
  return first.pages.length === second.pages.length && first.pages.every((page, index) => {
    const next = second.pages[index]!;
    return sameShellValue(page.page, next.page) && sameShellValue(page.bounds, next.bounds) &&
      sameShellValue(page.transform, next.transform);
  });
}

/** @internal Compares the view fields the React shell observes, excluding paint epochs and page content. */
export function sameShellView(first: PeerlessViewSnapshot | null, second: PeerlessViewSnapshot | null): boolean {
  return first === second || !!first && !!second && sameShellValue(first.identity, second.identity) &&
    first.version === second.version && first.throughSeq === second.throughSeq &&
    first.versionSinceSeq === second.versionSinceSeq && first.presentedThroughSeq === second.presentedThroughSeq &&
    sameShellValue(first.workerSelection, second.workerSelection) && sameShellValue(first.anchors, second.anchors) &&
    sameShellValue(first.editState, second.editState) && sameShellValue(first.provisionalCaret, second.provisionalCaret) &&
    sameGeometry(first.geometry, second.geometry);
}
