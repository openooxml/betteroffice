/**
 * Declaration-only bridge for the Vue adapter's retired core subpath imports.
 * Vue still owns its ProseMirror migration separately; keep these narrow
 * declarations until that adapter stops resolving the old paths from source.
 */
declare module '@betteroffice/docx/prosemirror/extensions/types' {
  export type CommandMap = Record<string, (...args: any[]) => any>;
}
