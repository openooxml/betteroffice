/** Engine operations the package runs on sessions it creates itself; not public API. */

export interface YrsSessionInternals {
  /** Compares two DOCX packages into the empty session; bridge JSON. */
  compareDocx(original: Uint8Array, revised: Uint8Array, options: string): string;
  /** The final comparison result for the saved bytes; bridge JSON. */
  finishComparedDocx(bytes: Uint8Array): string;
  /** The final comparison result when saving failed; bridge JSON. */
  failComparedDocx(message: string): string;
}

const registry = new WeakMap<object, YrsSessionInternals>();

export function registerSessionInternals(session: object, internals: YrsSessionInternals): void {
  registry.set(session, internals);
}

export function sessionInternals(session: object): YrsSessionInternals {
  const internals = registry.get(session);
  if (!internals) throw new Error('the session was not created by createYrsSession');
  return internals;
}
