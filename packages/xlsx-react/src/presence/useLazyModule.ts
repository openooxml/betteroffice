import { useEffect, useState } from 'react';

export function useLazyModule<T>(load: (() => Promise<T>) | null, retryKey: unknown): T | null {
  const [loaded, setLoaded] = useState<{ module: T } | null>(null);
  useEffect(() => {
    if (!load || loaded) return;
    let live = true;
    load().then((module) => { if (live) setLoaded({ module }); }, () => {});
    return () => { live = false; };
  }, [load, retryKey, loaded]);
  return load ? loaded?.module ?? null : null;
}
