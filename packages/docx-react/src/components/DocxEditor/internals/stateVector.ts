function decodeStateVector(bytes: Uint8Array): Map<number, number> {
  let at = 0;
  const read = (): number => {
    let value = 0;
    for (let scale = 1; ; scale *= 0x80) {
      const byte = bytes[at++];
      if (byte === undefined) throw new Error('Truncated yrs state vector');
      value += (byte & 0x7f) * scale;
      if (byte < 0x80) return value;
    }
  };
  const clocks = new Map<number, number>();
  for (let count = read(); count > 0; count--) {
    const client = read();
    clocks.set(client, read());
  }
  return clocks;
}

/** Whether `local` holds a client clock past `remote`'s, i.e. content `remote` lacks. */
export function stateVectorAhead(local: Uint8Array, remote: Uint8Array | null): boolean {
  if (!remote) return true;
  const known = decodeStateVector(remote);
  for (const [client, clock] of decodeStateVector(local)) {
    if (clock > (known.get(client) ?? 0)) return true;
  }
  return false;
}
