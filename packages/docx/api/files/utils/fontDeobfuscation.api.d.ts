// @betteroffice/docx declaration utils/fontDeobfuscation.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
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
/**
 * Whether a string is a usable embedded-font obfuscation key (a 128-bit GUID,
 * with or without braces/hyphens).
 *
 * @public
 */
export declare function isValidFontKey(fontKey: string | undefined | null): boolean;
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
export declare function deobfuscateFont(data: ArrayBuffer, fontKey: string): ArrayBuffer;
