import { useEffect, type RefObject } from 'react';
import { trackPluginGestures } from './gesture';

export function usePluginGestures(root: RefObject<HTMLElement | null>): void {
  useEffect(() => {
    if (root.current) return trackPluginGestures(root.current);
  }, [root]);
}
