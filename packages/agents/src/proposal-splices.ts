import type { ProposalChange } from './types';

export const MAX_SPLICE_CELLS = 1_000_000;

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

export function boundary(text: string, offset: number): boolean {
  return offset === 0 || offset === text.length ||
    !(text.charCodeAt(offset - 1) >= 0xd800 && text.charCodeAt(offset - 1) <= 0xdbff &&
      text.charCodeAt(offset) >= 0xdc00 && text.charCodeAt(offset) <= 0xdfff);
}

function graphemeCuts(text: string): Set<number> {
  const cuts = new Set<number>();
  for (const { index } of graphemes.segment(text)) if (boundary(text, index)) cuts.add(index);
  cuts.add(text.length);
  return cuts;
}

function tokenOffsets(text: string, cuts: Set<number>, from: number, to: number): number[] {
  const offsets = [from];
  let start = from;
  let previous: 'word' | 'space' | undefined;
  for (const end of cuts) {
    if (end <= from) continue;
    if (end > to) break;
    const cluster = text.slice(start, end);
    const kind = /^[\p{L}\p{N}\p{M}_]/u.test(cluster) ? 'word' : /^\s/u.test(cluster) ? 'space' : undefined;
    if (start !== from && (!kind || kind !== previous)) offsets.push(start);
    start = end;
    previous = kind;
  }
  if (to !== from) offsets.push(to);
  return offsets;
}

export function spliceChange(change: ProposalChange): ProposalChange[] {
  const { oldText, newText } = change;
  if (oldText === newText) return [];
  const oldCuts = graphemeCuts(oldText), newCuts = graphemeCuts(newText);
  let left = 0;
  while (left < oldText.length && left < newText.length && oldText[left] === newText[left]) left++;
  while (!(oldCuts.has(left) && newCuts.has(left))) left--;
  let right = 0;
  while (right < oldText.length - left && right < newText.length - left && oldText[oldText.length - right - 1] === newText[newText.length - right - 1]) right++;
  while (!(oldCuts.has(oldText.length - right) && newCuts.has(newText.length - right))) right--;
  const oldOffsets = tokenOffsets(oldText, oldCuts, left, oldText.length - right);
  const newOffsets = tokenOffsets(newText, newCuts, left, newText.length - right);
  const oldCount = oldOffsets.length - 1, newCount = newOffsets.length - 1;
  const splice = (oldFrom: number, oldTo: number, newFrom: number, newTo: number): ProposalChange => ({
    ...change, start: change.start + oldOffsets[oldFrom]!,
    oldText: oldText.slice(oldOffsets[oldFrom], oldOffsets[oldTo]),
    newText: newText.slice(newOffsets[newFrom], newOffsets[newTo]),
  });
  const columns = newCount + 1;
  const cells = (oldCount + 1) * columns;
  if (!oldCount || !newCount || cells > MAX_SPLICE_CELLS) return [splice(0, oldCount, 0, newCount)];

  const oldTokens = oldOffsets.slice(0, -1).map((at, i) => oldText.slice(at, oldOffsets[i + 1]));
  const newTokens = newOffsets.slice(0, -1).map((at, j) => newText.slice(at, newOffsets[j + 1]));
  const lcs = new Uint32Array(cells);
  for (let i = oldCount - 1; i >= 0; i--) {
    const row = i * columns, next = row + columns;
    for (let j = newCount - 1; j >= 0; j--) {
      lcs[row + j] = oldTokens[i] === newTokens[j]
        ? 1 + lcs[next + j + 1]!
        : Math.max(lcs[next + j]!, lcs[row + j + 1]!);
    }
  }

  const splices: ProposalChange[] = [];
  let oldFrom = 0, newFrom = 0;
  for (let i = 0; i < oldCount; i++) {
    const row = i * columns;
    const remaining = lcs[row + newFrom]!;
    if (!remaining) break;
    for (let j = newFrom; j < newCount && lcs[row + j] === remaining; j++) {
      if (oldTokens[i] !== newTokens[j]) continue;
      if (i !== oldFrom || j !== newFrom) splices.push(splice(oldFrom, i, newFrom, j));
      oldFrom = i + 1;
      newFrom = j + 1;
      break;
    }
  }
  if (oldFrom !== oldCount || newFrom !== newCount) splices.push(splice(oldFrom, oldCount, newFrom, newCount));
  return splices;
}
