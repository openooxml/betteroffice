/**
 * Tiny shared CSSProperties shape so framework-agnostic style
 * objects in core don't need to import from React or Vue. Both
 * adapters' `CSSProperties` types are structurally compatible
 * with this looser shape (string | number values, camelCased keys).
 */
export type CSSProperties = Record<string, string | number | undefined>;
