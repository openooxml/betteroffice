/** @internal The message a native engine call fails with once its thread is past its memory budget. */
export const NATIVE_MEMORY_BUDGET_EXCEEDED = 'BetterOffice engine memory budget exceeded';

/** @internal Whether `error` leaves the engine unusable: a WebAssembly trap, or a native call that ran past its memory budget. */
export function isEngineTrap(error: unknown): boolean {
  return (typeof WebAssembly !== 'undefined' && error instanceof WebAssembly.RuntimeError) ||
    (error instanceof Error && error.message === NATIVE_MEMORY_BUDGET_EXCEEDED);
}
