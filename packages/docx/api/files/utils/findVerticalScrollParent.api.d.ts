// @betteroffice/docx declaration utils/findVerticalScrollParent.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Pure DOM helpers — locate the element that vertically scrolls the
 * paginated editor. Lifted from packages/docx-react/src/paged-editor/
 * findVerticalScrollParent.ts so both adapters use the same logic
 * for scroll-to-position, arrow-key line navigation, and drag
 * auto-scroll.
 * @packageDocumentation
 * @public
 */
/**
 * First ancestor of `el` with `overflow-y: auto|scroll` and a scrollable
 * overflow height. Walk starts at `el.parentElement` (does not treat
 * `el` itself as the scroller).
 *
 * @returns `null` when no such ancestor exists before `document.documentElement`.
 */
export declare function findVerticalScrollParent(el: HTMLElement): HTMLElement | null;
/**
 * Same as {@link findVerticalScrollParent} but falls back to
 * `document.documentElement` so callers always get a valid scroll
 * target (matches legacy `scrollIntoView` root).
 */
export declare function findVerticalScrollParentOrRoot(el: HTMLElement): HTMLElement;
