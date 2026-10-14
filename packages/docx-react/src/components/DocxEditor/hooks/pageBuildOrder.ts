/** The `count` pages of ascending `unbuilt` (none in `[start, end)`) nearest the window, ascending; ties go to the lower page. */
export function nearestPages(
  unbuilt: readonly number[],
  start: number,
  end: number,
  count: number
): number[] {
  if (count <= 0 || unbuilt.length === 0) return [];

  let low = 0;
  let high = unbuilt.length;
  while (low < high) {
    const middle = low + Math.floor((high - low) / 2);
    if (unbuilt[middle] < start) low = middle + 1;
    else high = middle;
  }

  let left = low - 1;
  let right = low;
  const limit = Math.min(count, unbuilt.length);
  for (let selected = 0; selected < limit; selected += 1) {
    if (
      right === unbuilt.length ||
      (left >= 0 && start - unbuilt[left] <= unbuilt[right] - end + 1)
    ) {
      left -= 1;
    } else {
      right += 1;
    }
  }
  return unbuilt.slice(left + 1, right);
}
