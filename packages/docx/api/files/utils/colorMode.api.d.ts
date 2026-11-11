// @betteroffice/docx declaration utils/colorMode.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Color-mode resolution shared by every adapter so the light/dark/auto logic
 * (and its SSR + prefers-color-scheme handling) can't drift between React, Vue,
 * or any future host. The adapters keep only their framework binding
 * (useState/useEffect vs ref/watchEffect); the rules live here.
 *
 * @packageDocumentation
 */
/** UI color theme for the editor chrome and canvas. */
export type ColorMode = 'light' | 'dark' | 'system';
/**
 * Current OS dark-mode preference. Returns false when `matchMedia` is
 * unavailable (e.g. server-side render), so it is safe to call as a state seed.
 *
 * @public
 */
export declare function prefersColorSchemeDark(): boolean;
/**
 * Resolve the effective dark flag from a {@link ColorMode} and the current OS
 * preference. `'system'` follows the OS; `'dark'`/`'light'` are explicit.
 *
 * @public
 */
export declare function resolveIsDark(colorMode: ColorMode, systemDark: boolean): boolean;
/**
 * Subscribe to OS dark-mode changes. Invokes `onChange` immediately with the
 * current value (so a stale seed is corrected on entry), then on every change.
 * Returns an unsubscribe function. A no-op under SSR.
 *
 * @public
 */
export declare function subscribeSystemDark(onChange: (dark: boolean) => void): () => void;
