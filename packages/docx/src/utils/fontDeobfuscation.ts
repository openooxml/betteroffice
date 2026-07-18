/**
 * Embedded font de-obfuscation (ECMA-376 Part 4 §2.8.1, "Font Embedding").
 *
 * Word stores embedded fonts as obfuscated OpenType (`.odttf`): the first 32
 * bytes of the font binary are XOR-scrambled with a 16-byte key, the rest of
 * the file is untouched. The key is the font's `w:fontKey` GUID with its byte
 * order REVERSED, applied once to bytes 0-15 and again to bytes 16-31.
 *
 * Example (from the spec): GUID `001B70DC-AA60-4AD5-90EC-18A0948E1EAE` yields
 * the key bytes `AE 1E 8E 94 A0 18 EC 90 D5 4A 60 AA DC 70 1B 00`.
 *
 * The scheme is a pure XOR, so the same operation obfuscates and de-obfuscates.
 */

const HEADER_LENGTH = 32;
const KEY_LENGTH = 16;

/**
 * Strip a `w:fontKey` GUID down to its 16 key bytes, in the reversed order the
 * obfuscation applies. Returns null when the value is not a 32-hex-digit GUID.
 */
function fontKeyToReversedBytes(fontKey: string): Uint8Array | null {
  const hex = fontKey.replace(/[^0-9a-fA-F]/g, '');
  if (hex.length !== KEY_LENGTH * 2) return null;

  const forward = new Uint8Array(KEY_LENGTH);
  for (let i = 0; i < KEY_LENGTH; i++) {
    const byte = Number.parseInt(hex.slice(i * 2, i * 2 + 2), 16);
    if (Number.isNaN(byte)) return null;
    forward[i] = byte;
  }

  // Reverse the byte order — this is the actual XOR key.
  const reversed = new Uint8Array(KEY_LENGTH);
  for (let i = 0; i < KEY_LENGTH; i++) {
    reversed[i] = forward[KEY_LENGTH - 1 - i];
  }
  return reversed;
}

/**
 * Whether a string is a usable embedded-font obfuscation key (a 128-bit GUID,
 * with or without braces/hyphens).
 *
 * @public
 */
export function isValidFontKey(fontKey: string | undefined | null): boolean {
  if (!fontKey) return false;
  return fontKeyToReversedBytes(fontKey) !== null;
}

/**
 * De-obfuscate an embedded `.odttf` font into a usable OpenType/TrueType
 * binary by XOR-ing its first 32 bytes with the reversed `w:fontKey` GUID.
 *
 * Returns a new buffer; the input is not mutated. Throws if `fontKey` is not a
 * valid 128-bit GUID.
 *
 * @param data - Raw obfuscated font bytes from `word/fonts/*.odttf`.
 * @param fontKey - The `w:fontKey` GUID from the `w:embed*` element.
 * @public
 */
export function deobfuscateFont(data: ArrayBuffer, fontKey: string): ArrayBuffer {
  const key = fontKeyToReversedBytes(fontKey);
  if (!key) {
    throw new Error(`Invalid embedded-font key: "${fontKey}"`);
  }

  const out = new Uint8Array(data.slice(0));
  const end = Math.min(HEADER_LENGTH, out.length);
  for (let i = 0; i < end; i++) {
    out[i] ^= key[i % KEY_LENGTH];
  }
  return out.buffer;
}
