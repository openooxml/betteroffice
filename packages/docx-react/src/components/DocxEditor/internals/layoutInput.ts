function canonical(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonical);
  if (value === null || typeof value !== 'object') return value;
  return Object.fromEntries(
    Object.keys(value)
      .sort()
      .map((key) => [key, canonical((value as Record<string, unknown>)[key])])
  );
}

export function sameLayoutInput(held: string, request: string): boolean {
  if (held === request) return true;
  try {
    return JSON.stringify(canonical(JSON.parse(held))) === JSON.stringify(canonical(JSON.parse(request)));
  } catch {
    return false;
  }
}
