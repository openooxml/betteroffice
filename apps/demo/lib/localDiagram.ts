/** What the demo needs of a picked file; a DOM `File` satisfies it. */
export interface PickedFile {
  name: string;
  arrayBuffer(): Promise<ArrayBuffer>;
}

export type LocalDiagram =
  | { opened: { name: string; bytes: Uint8Array } }
  | { refused: string };

/**
 * Reads a picked file and lets `accept` refuse it, so only a file the engine
 * can open replaces the loaded diagram.
 */
export async function readLocalDiagram(
  picked: PickedFile,
  accept: (bytes: Uint8Array) => Promise<void>,
): Promise<LocalDiagram> {
  try {
    const bytes = new Uint8Array(await picked.arrayBuffer());
    await accept(bytes);
    return { opened: { name: picked.name, bytes } };
  } catch (cause) {
    const reason = cause instanceof Error ? cause.message : String(cause);
    return { refused: `Could not open “${picked.name}”: ${reason} Previous diagram kept.` };
  }
}
