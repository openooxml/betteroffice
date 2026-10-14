/**
 * @internal The pages of a cut preview's layout that the whole document lays
 * out the same, or null for a whole document. A prefix pass fixes its first
 * `provisionalPages`; a pass that reached the cut, all but its last two.
 */
export function finalPreviewPageCount(
  partialDocument: boolean | undefined,
  provisional: boolean,
  provisionalPages: number | undefined,
  pageCount: number | null
): number | null {
  if (partialDocument !== true) return null;
  if (provisional) return provisionalPages ?? 0;
  return pageCount === null ? 0 : Math.max(0, pageCount - 2);
}

/** @internal `window` limited to a cut preview's first `finalPages` pages. */
export function finalPreviewDisplayWindow(
  window: [number, number] | undefined,
  finalPages: number
): [number, number] {
  const [start, end] = window ?? [0, finalPages];
  return [start, Math.min(end, Math.max(start, finalPages))];
}
