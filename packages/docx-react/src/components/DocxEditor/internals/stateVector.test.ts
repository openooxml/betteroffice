import { expect, test } from 'bun:test';
import { stateVectorAhead } from './stateVector';

test('a null remote state vector needs the local state', () => {
  expect(stateVectorAhead(Uint8Array.of(1, 3, 5), null)).toBe(true);
});

test('equal state vectors are not ahead regardless of client order', () => {
  expect(stateVectorAhead(Uint8Array.of(2, 3, 5, 7, 9), Uint8Array.of(2, 7, 9, 3, 5))).toBe(false);
});

test('a local state vector behind the remote is not ahead', () => {
  expect(stateVectorAhead(Uint8Array.of(1, 3, 4), Uint8Array.of(2, 3, 5, 7, 9))).toBe(false);
});

test('one local client clock ahead requires a diff', () => {
  expect(stateVectorAhead(Uint8Array.of(2, 3, 6, 7, 8), Uint8Array.of(2, 3, 5, 7, 9))).toBe(true);
});

test('a local client unknown to the remote requires a diff', () => {
  expect(stateVectorAhead(Uint8Array.of(2, 3, 5, 7, 1), Uint8Array.of(1, 3, 5))).toBe(true);
});

test('multi-byte client IDs above 2^32 and clocks above 127 compare without truncation', () => {
  const local = Uint8Array.of(1, 0x81, 0x80, 0x80, 0x80, 0x10, 0x80, 0x02);
  const behind = Uint8Array.of(1, 0x81, 0x80, 0x80, 0x80, 0x10, 0xff, 0x01);
  expect(stateVectorAhead(local, local.slice())).toBe(false);
  expect(stateVectorAhead(local, behind)).toBe(true);
  expect(stateVectorAhead(behind, local)).toBe(false);
  expect(stateVectorAhead(local, Uint8Array.of(1, 1, 0x80, 0x02))).toBe(true);
});
