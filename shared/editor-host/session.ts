/** Saved bytes, or the failure the editor reported, or that it is opening. */
export function savedBytes(
  bytes: ArrayBuffer | Uint8Array | null | undefined,
  failure: Error | null,
  stillOpening = "The editor is still opening the file. Try again in a moment."
): Uint8Array {
  if (bytes) return new Uint8Array(bytes);
  throw failure ?? new Error(stillOpening);
}
