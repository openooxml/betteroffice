import { SessionFailure } from '../types';

const MAX_LEDGER_RANGES = 4096;

/** @internal Retains blocked sequences in bounded ranges for one document epoch. */
export function createBlockedLedger() {
  const ranges: [number, number][] = [];

  function contains(seq: number): boolean {
    let low = 0;
    let high = ranges.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      const range = ranges[middle]!;
      if (seq < range[0]) high = middle;
      else if (seq > range[1]) low = middle + 1;
      else return true;
    }
    return false;
  }

  function add(seq: number): void {
    let index = 0;
    while (index < ranges.length && ranges[index]![1] < seq - 1) index += 1;
    const next = ranges[index];
    if (next && next[0] <= seq + 1) {
      next[0] = Math.min(next[0], seq);
      next[1] = Math.max(next[1], seq);
      const following = ranges[index + 1];
      if (following && following[0] <= next[1] + 1) {
        next[1] = Math.max(next[1], following[1]);
        ranges.splice(index + 1, 1);
      }
      return;
    }
    if (ranges.length === MAX_LEDGER_RANGES) {
      throw new SessionFailure('out-of-memory', 'Command replay ledger exhausted: 4096 disjoint ranges per epoch');
    }
    ranges.splice(index, 0, [seq, seq]);
  }

  return {
    contains, add,
    values() {
      const seqs: number[] = [];
      for (const [start, end] of ranges) {
        for (let seq = start; seq <= end; seq += 1) seqs.push(seq);
      }
      return seqs;
    },
    clear() { ranges.length = 0; },
  };
}
