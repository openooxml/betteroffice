export class SessionRevision {
  private revision = 0;
  private saved = 0;

  get dirty(): boolean {
    return this.revision !== this.saved;
  }
  change(): void {
    this.revision += 1;
  }
  checkpoint(): number {
    return this.revision;
  }
  didSave(checkpoint: number): void {
    this.saved = checkpoint;
  }
}

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
