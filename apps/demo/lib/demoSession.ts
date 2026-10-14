/**
 * The loaded document. `seed` is the room state it hydrates from; a locally
 * opened file has none. The relay never retains a first peer's document, so
 * such a file cannot be rejoined or reloaded and is therefore not shared.
 */
export interface DemoDocument {
  seed: Uint8Array | null;
}

export interface DemoSessionInputs {
  document: DemoDocument | null;
  room: string | null;
  clientId: number | null;
  identified: boolean;
}

/** `loading` withholds the editor entirely; only `shared` may collaborate. */
export type DemoSession =
  | { status: 'loading' }
  | { status: 'local' }
  | { status: 'shared'; room: string; clientId: number };

/**
 * The one gate deciding how the editor mounts. A seeded document waits for its
 * whole room rather than mounting unshared, which would seed the room's
 * document into an independent replica.
 */
export function planDemoSession({
  document,
  room,
  clientId,
  identified,
}: DemoSessionInputs): DemoSession {
  if (!document) return { status: 'loading' };
  if (document.seed === null) return { status: 'local' };
  if (!room || clientId === null || !identified) return { status: 'loading' };
  return { status: 'shared', room, clientId };
}
