/**
 * Cleanup that never replaces the failure it runs after. `failed` cleans up and
 * returns that failure, recording a cleanup error on it as `suppressed`; `run`
 * cleans up after success, where a cleanup error is the only one to report.
 */
export function cleanupGuard(cleanup: () => void) {
  let done = false;
  return {
    failed<T>(failure: T): T {
      if (done) return failure;
      done = true;
      try { cleanup(); }
      catch (error) {
        if (failure instanceof Error) {
          try { Reflect.defineProperty(failure, 'suppressed', { value: error, configurable: true, writable: true }); }
          catch { /* the failure stays the one to report */ }
        }
      }
      return failure;
    },
    run() {
      if (done) return;
      done = true;
      cleanup();
    },
  };
}
