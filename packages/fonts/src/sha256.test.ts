import { expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { sha256, sha256Fallback } from './sha256';

for (const length of [0, 3, 55, 56, 63, 64, 65, 4096]) {
  test(`hashes ${length} bytes with and without WebCrypto`, async () => {
    const bytes = Uint8Array.from({ length }, (_, index) => index & 255);
    const expected = createHash('sha256').update(bytes).digest('hex');
    expect(sha256Fallback(bytes.buffer)).toBe(expected);
    expect(await sha256(bytes.buffer)).toBe(expected);
  });
}
