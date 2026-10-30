export const BASELINE_GENERATION_HEADER = "X-WOPI-Baseline-Generation";
export const BASELINE_GENERATION_PROTOCOL = "bo-wopi.generation.";

export function parseBaselineGeneration(value: string | null): number | null {
  if (value === null || !/^(0|[1-9]\d*)$/.test(value)) return null;
  const generation = Number(value);
  return Number.isSafeInteger(generation) ? generation : null;
}
