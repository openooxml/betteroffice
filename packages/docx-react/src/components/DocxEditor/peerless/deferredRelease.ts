export interface DeferredRelease {
  acquire(): boolean;
  release(): void;
}

/** Disposes once, one task after the last release, unless acquired again first; acquire reports whether it is still undisposed. */
export function createDeferredRelease(dispose: () => void): DeferredRelease {
  let timer: ReturnType<typeof setTimeout> | null = null;
  let disposed = false;
  return {
    acquire() {
      if (timer !== null) clearTimeout(timer);
      timer = null;
      return !disposed;
    },
    release() {
      if (disposed || timer !== null) return;
      timer = setTimeout(() => {
        timer = null;
        disposed = true;
        dispose();
      }, 0);
    },
  };
}
