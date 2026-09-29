import { useEffect, useRef, useState } from 'react';
import {
  createFontLoadScope,
  type FontDefinition,
  type FontLoadScope,
} from '@betteroffice/docx/utils';

/**
 * The editor instance's font load scope: its loads notify only its own
 * listeners, and its embedded faces are released on unmount. A StrictMode
 * remount revives the scope its cleanup disposed before the effects declared
 * after this hook, such as document loads, run again.
 */
export function useFontLoadScope(): FontLoadScope {
  const [scope] = useState(createFontLoadScope);
  useEffect(() => {
    const unsubscribe = scope.onFontsLoaded(() => {});
    return () => {
      unsubscribe();
      scope.dispose();
    };
  }, [scope]);
  return scope;
}

/**
 * Owns the editor's three font lifecycle wires:
 *
 * 1. Re-register custom faces from the `fonts` prop on identity change.
 *    The loader dedupes by `family|weight`, so re-runs are cheap.
 * 2. Forward this instance's `onFontsLoaded` events to the consumer's callback.
 * 3. Forward this instance's font-load failures to the consumer's `onError`
 *    prop. The subscription reads `onError` through a ref so an inline
 *    `onError={(e) => …}` does not churn the subscriber Set on every parent
 *    render.
 *
 * Loads another editor instance starts do not reach this one; module-level
 * loads still do.
 */
export function useFontLifecycle(
  fonts: ReadonlyArray<FontDefinition> | undefined,
  onFontsLoadedCallback: (() => void) | undefined,
  onError: ((error: Error) => void) | undefined,
  scope: FontLoadScope
): void {
  useEffect(() => {
    void scope.loadFontDefinitions(fonts);
  }, [fonts, scope]);

  useEffect(() => {
    return scope.onFontsLoaded(() => onFontsLoadedCallback?.());
  }, [onFontsLoadedCallback, scope]);

  const onErrorRef = useRef(onError);
  useEffect(() => {
    onErrorRef.current = onError;
  }, [onError]);
  useEffect(() => {
    return scope.onFontError((err) => onErrorRef.current?.(err));
  }, [scope]);
}
