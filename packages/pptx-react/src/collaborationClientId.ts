export function generateCollaborationClientId(): number {
  const values = new Uint32Array(1);
  do globalThis.crypto.getRandomValues(values); while (values[0] === 0);
  return values[0]!;
}
