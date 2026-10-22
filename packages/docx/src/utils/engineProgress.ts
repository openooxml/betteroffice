const TICKS_PER_REPORT = 256;
let ticks = 0;

/** Tells a worker host that a long engine call is still running; a no-op where none listens. @internal */
export function engineProgress(): void {
  (globalThis as { __betterofficeEngineProgress?: () => void }).__betterofficeEngineProgress?.();
}

/** {@link engineProgress} once every few hundred calls, for per-item loops. @internal */
export function engineProgressTick(): void {
  ticks += 1;
  if (ticks % TICKS_PER_REPORT === 0) engineProgress();
}
