/** Saved bytes, or the failure the editor reported, or that it is opening. */
export function savedBytes(
  bytes: ArrayBuffer | Uint8Array | null | undefined,
  failure: Error | null
): Uint8Array {
  if (bytes) return new Uint8Array(bytes);
  throw (
    failure ??
    new Error("The editor is still opening the file. Try again in a moment.")
  );
}
