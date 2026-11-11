/**
 * Gives native engine classes the `[Symbol.dispose]` the wasm glue defines:
 * the same function as `free`, assigned where the runtime has `Symbol.dispose`.
 */
export function installDispose(...classes: ReadonlyArray<{ prototype: { free(): void } }>): void {
  if (!Symbol.dispose) return;
  for (const engineClass of classes) {
    (engineClass.prototype as Record<symbol, unknown>)[Symbol.dispose] = engineClass.prototype.free;
  }
}
