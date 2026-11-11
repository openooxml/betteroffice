/**
 * Whether `error` is the engine's refusal of a write into a table cell whose story is missing.
 * The engine refuses it before any mutation; callers still check that nothing was written.
 * @internal
 */
export function isCellHasNoStoryRefusal(error: unknown): boolean {
  return error instanceof Error && 'code' in error && error.code === 'cell-has-no-story';
}
