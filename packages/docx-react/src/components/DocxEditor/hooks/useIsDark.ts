import { useEffect, useState } from 'react';
import { prefersColorSchemeDark, resolveIsDark, subscribeSystemDark } from '@betteroffice/docx/utils';

// resolves the effective dark flag for a colorMode prop, tracking the OS
// scheme while colorMode is 'system'. subscribeSystemDark re-syncs
// immediately (correcting a stale seed if the OS theme changed while
// colorMode was 'light'/'dark') and is SSR-safe.
export function useIsDark(colorMode: 'light' | 'dark' | 'system'): boolean {
  const [systemDark, setSystemDark] = useState(prefersColorSchemeDark);
  useEffect(() => {
    if (colorMode !== 'system') return;
    return subscribeSystemDark(setSystemDark);
  }, [colorMode]);
  return resolveIsDark(colorMode, systemDark);
}
