import { describe, expect, it } from 'bun:test';
import { isClientMessage, isHostMessage } from './protocol';

describe('wasm compile protocol', () => {
  it('accepts compile messages only from the client', () => {
    const message = { protocol: 1, kind: 'wasm-compile' };
    expect(isClientMessage(message)).toBe(true);
    expect(isHostMessage(message)).toBe(false);
  });

  it('rejects malformed compile messages', () => {
    for (const message of [
      null, undefined, 'wasm-compile', [],
      { kind: 'wasm-compile' },
      { protocol: 0, kind: 'wasm-compile' },
      { protocol: 2, kind: 'wasm-compile' },
      { protocol: '1', kind: 'wasm-compile' },
      { protocol: 1 },
      { protocol: 1, kind: null },
      { protocol: 1, kind: 'wasm_compile' },
    ]) {
      expect(isClientMessage(message)).toBe(false);
    }
  });
});
