import { expect, test } from 'bun:test';
import { SessionFailure } from '../types';
import { createBlockedLedger } from './ledger';

test('merges adjacent blocked ranges and bridges them at the epoch cap', () => {
  const ledger = createBlockedLedger();
  for (let seq = 1; seq <= 8191; seq += 2) ledger.add(seq);
  ledger.add(8192);
  ledger.add(2);
  ledger.add(8194);
  expect(ledger.contains(2)).toBe(true);
  expect(ledger.contains(8192)).toBe(true);
  expect(ledger.contains(8193)).toBe(false);
  expect(ledger.contains(8194)).toBe(true);
  ledger.add(2);
  expect(ledger.values().slice(0, 3)).toEqual([1, 2, 3]);
});

test('reports explicit terminal exhaustion at 4096 disjoint blocked ranges per epoch', () => {
  const ledger = createBlockedLedger();
  for (let seq = 1; seq <= 8191; seq += 2) ledger.add(seq);
  let failure: unknown;
  try { ledger.add(8193); } catch (error) { failure = error; }
  expect(failure).toBeInstanceOf(SessionFailure);
  expect(failure).toMatchObject({
    code: 'out-of-memory', message: 'Command replay ledger exhausted: 4096 disjoint ranges per epoch',
  });
  expect(ledger.contains(8193)).toBe(false);
  expect(ledger.contains(8191)).toBe(true);
  ledger.clear();
  ledger.add(8193);
  expect(ledger.values()).toEqual([8193]);
});
