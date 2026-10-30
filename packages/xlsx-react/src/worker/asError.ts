export function asError(value: unknown): Error {
  return value instanceof Error ? value : new Error(
    value && typeof value === 'object' && 'message' in value ? String(value.message) : String(value)
  );
}
