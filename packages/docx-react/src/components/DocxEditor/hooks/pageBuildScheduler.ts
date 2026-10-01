const FRAME_BUDGET_MS = 8;

export type PageBuildTask = { cancel(): void };

export function scheduleIdlePageBuild(
  run: (deadline: Pick<IdleDeadline, 'timeRemaining'>) => void
): PageBuildTask {
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
  const id = requestAnimationFrame((frameStart) => {
    const timeRemaining = () => Math.max(0, FRAME_BUDGET_MS - (performance.now() - frameStart));
    if (timeRemaining() > 0) run({ timeRemaining });
    else next = scheduleIdlePageBuild(run);
  });
  return {
    cancel() {
      cancelAnimationFrame(id);
      next?.cancel();
    },
  };
}
