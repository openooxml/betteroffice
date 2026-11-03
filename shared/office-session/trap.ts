/** @internal The message a native engine call fails with once its thread is past its memory budget. */
export const NATIVE_MEMORY_BUDGET_EXCEEDED = 'BetterOffice engine memory budget exceeded';

/** @internal The message of a native engine call refused because an earlier call on the same object trapped. */
export const NATIVE_ENGINE_OBJECT_FAILED = 'This engine object failed in an earlier call and can only be freed';

/**
 * @internal Whether `error` leaves the engine unusable: a WebAssembly trap, a native call that ran
 * past its memory budget, or a native call refused because an earlier call on its object trapped.
 */
export function isEngineTrap(error: unknown): boolean {
  return (typeof WebAssembly !== 'undefined' && error instanceof WebAssembly.RuntimeError) ||
    (error instanceof Error &&
      (error.message.startsWith(NATIVE_MEMORY_BUDGET_EXCEEDED) || error.message === NATIVE_ENGINE_OBJECT_FAILED));
}

const NATIVE_BUDGET_DETAIL = /BetterOffice engine memory budget exceeded \((\d+) bytes over a (\d+)-byte budget\)/;

/** @internal The allocation and budget a native engine error reports, found anywhere in `text`. */
export function nativeBudgetFailure(text: string | undefined): { bytes: number; budgetBytes: number } | null {
  const match = text === undefined ? null : NATIVE_BUDGET_DETAIL.exec(text);
  return match ? { bytes: Number(match[1]), budgetBytes: Number(match[2]) } : null;
}
