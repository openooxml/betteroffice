import type { ProposalChange } from './types';

export const MAX_SPLICE_CELLS = 1_000_000;
export const MAX_PROPOSAL_STEPS = 128;

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
  const edges = [...cuts];
  const kinds = edges.slice(0, -1).map((start, i) => {
    const cluster = text.slice(start, edges[i + 1]);
    if (/^[\p{L}\p{N}\p{M}]/u.test(cluster)) return 'word';
    if (/^\s/u.test(cluster)) return 'space';
    if (/^[.,/:'’_@-]$/u.test(cluster)) return 'joiner';
  });
  const offsets = [from];
  let previous: 'word' | 'space' | undefined;
  for (const [i, start] of edges.entries()) {
    if (start < from) continue;
    if (start >= to) break;
    const cluster = kinds[i];
    const kind = cluster === 'joiner'
      ? (kinds[i - 1] === 'word' && kinds[i + 1] === 'word' ? 'word' : undefined)
      : cluster;
    if (start !== from && (!kind || kind !== previous)) offsets.push(start);
    previous = kind;
  }
  if (to !== from) offsets.push(to);
  return offsets;
}

export function trimChange(change: ProposalChange, oldCuts = graphemeCuts(change.oldText), newCuts = graphemeCuts(change.newText)): ProposalChange {
  const { oldText, newText } = change;
  let left = 0;
  while (left < oldText.length && left < newText.length && oldText[left] === newText[left]) left++;
  while (!(oldCuts.has(left) && newCuts.has(left))) left--;
  let right = 0;
  while (right < oldText.length - left && right < newText.length - left && oldText[oldText.length - right - 1] === newText[newText.length - right - 1]) right++;
  while (!(oldCuts.has(oldText.length - right) && newCuts.has(newText.length - right))) right--;
  return { ...change, start: change.start + left, oldText: oldText.slice(left, oldText.length - right), newText: newText.slice(left, newText.length - right) };
}

export function spliceChange(change: ProposalChange): ProposalChange[] {
  const { oldText, newText } = change;
  if (oldText === newText) return [];
  const oldCuts = graphemeCuts(oldText), newCuts = graphemeCuts(newText);
  const trimmed = trimChange(change, oldCuts, newCuts);
  if (!trimmed.oldText || !trimmed.newText) return [trimmed];
  const left = trimmed.start - change.start;
  const oldOffsets = tokenOffsets(oldText, oldCuts, left, left + trimmed.oldText.length);
  const newOffsets = tokenOffsets(newText, newCuts, left, left + trimmed.newText.length);
  const oldCount = oldOffsets.length - 1, newCount = newOffsets.length - 1;
  const splice = (oldFrom: number, oldTo: number, newFrom: number, newTo: number): ProposalChange => trimChange({
    ...change, start: change.start + oldOffsets[oldFrom]!,
    oldText: oldText.slice(oldOffsets[oldFrom], oldOffsets[oldTo]),
    newText: newText.slice(newOffsets[newFrom], newOffsets[newTo]),
  });
  const columns = newCount + 1;
  const cells = (oldCount + 1) * columns;
  if (cells > MAX_SPLICE_CELLS) return [trimmed];

  const oldTokens = oldOffsets.slice(0, -1).map((at, i) => oldText.slice(at, oldOffsets[i + 1]));
  const newTokens = newOffsets.slice(0, -1).map((at, j) => newText.slice(at, newOffsets[j + 1]));
  const weights = oldTokens.map(token => /^\s/u.test(token) ? 1 : token.length);
  const lcs = new Uint32Array(cells);
  for (let i = oldCount - 1; i >= 0; i--) {
    const row = i * columns, next = row + columns;
    for (let j = newCount - 1; j >= 0; j--) {
      lcs[row + j] = oldTokens[i] === newTokens[j]
        ? weights[i]! + lcs[next + j + 1]!
        : Math.max(lcs[next + j]!, lcs[row + j + 1]!);
    }
  }

  const splices: ProposalChange[] = [];
  const appendSplice = (next: ProposalChange) => {
    if (!next.oldText && !next.newText) return;
    const previous = splices.at(-1);
    if (previous && (!previous.oldText || !previous.newText || !next.oldText || !next.newText)) {
      const gap = oldText.slice(previous.start - change.start + previous.oldText.length, next.start - change.start);
      if (/^\s+$/u.test(gap)) {
        previous.oldText += gap + next.oldText;
        previous.newText += gap + next.newText;
        return;
      }
    }
    splices.push(next);
  };
  let oldFrom = 0, newFrom = 0;
  for (let i = 0; i < oldCount; i++) {
    const row = i * columns;
    const remaining = lcs[row + newFrom]!;
    if (!remaining) break;
    for (let j = newFrom; j < newCount && lcs[row + j] === remaining; j++) {
      if (oldTokens[i] !== newTokens[j]) continue;
      if (i !== oldFrom || j !== newFrom) appendSplice(splice(oldFrom, i, newFrom, j));
      oldFrom = i + 1;
      newFrom = j + 1;
      break;
    }
  }
  if (oldFrom !== oldCount || newFrom !== newCount) appendSplice(splice(oldFrom, oldCount, newFrom, newCount));
  return splices;
}
