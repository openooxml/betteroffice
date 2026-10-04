const FRAME_BUDGET_MS = 8;
const FRAME_FALLBACK_MS = 200;

export type PageBuildTask = { cancel(): void };

export function scheduleIdleWork(
  run: (deadline?: IdleDeadline) => void,
  expiresAt: number,
  fallbackDelay?: number
): PageBuildTask {
  const remaining = Math.max(0, expiresAt - performance.now());
  let idleId: number | null = null;
  let cancelled = false;
  const cancel = () => {
    cancelled = true;
    clearTimeout(timer);
    if (idleId !== null) {
      cancelIdleCallback(idleId);
      idleId = null;
    }
  };
  const finish = (deadline?: IdleDeadline) => {
    if (cancelled) return;
    cancel();
    run(performance.now() >= expiresAt ? undefined : deadline);
  };
  const hasIdle = typeof requestIdleCallback === 'function';
  const timer = setTimeout(
    () => finish(),
    hasIdle ? remaining : Math.min(fallbackDelay ?? remaining, remaining)
  );
  if (hasIdle && remaining > 0) {
    idleId = requestIdleCallback(finish, { timeout: remaining });
  }
  return { cancel };
}

export function scheduleIdlePageBuild(
  run: (deadline: Pick<IdleDeadline, 'timeRemaining'>) => void,
  urgent = false
): PageBuildTask {
  if (urgent) {
    const id = setTimeout(() => {
      const started = performance.now();
      run({ timeRemaining: () => Math.max(0, FRAME_BUDGET_MS - (performance.now() - started)) });
    }, 0);
    return { cancel: () => clearTimeout(id) };
  }
  if (typeof requestIdleCallback === 'function') {
    const id = requestIdleCallback((deadline) => {
      const started = performance.now();
      run({
        timeRemaining: () => Math.min(
          deadline?.timeRemaining() ?? FRAME_BUDGET_MS,
          Math.max(0, FRAME_BUDGET_MS - (performance.now() - started))
        ),
      });
    });
    return { cancel: () => cancelIdleCallback(id) };
  }
  let next: PageBuildTask | null = null;
  // Hidden tabs pause animation frames, so a timer runs the work if no frame comes.
  const timer = setTimeout(() => {
    cancelAnimationFrame(id);
    const started = performance.now();
    run({ timeRemaining: () => Math.max(0, FRAME_BUDGET_MS - (performance.now() - started)) });
  }, FRAME_FALLBACK_MS);
  const id = requestAnimationFrame((frameStart) => {
    clearTimeout(timer);
    const timeRemaining = () => Math.max(0, FRAME_BUDGET_MS - (performance.now() - frameStart));
    if (timeRemaining() > 0) run({ timeRemaining });
    else next = scheduleIdlePageBuild(run);
  });
  return {
    cancel() {
      clearTimeout(timer);
      cancelAnimationFrame(id);
      next?.cancel();
    },
  };
}
