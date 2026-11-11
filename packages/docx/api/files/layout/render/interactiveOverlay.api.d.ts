// @betteroffice/docx declaration layout/render/interactiveOverlay.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Framework-neutral interactive overlay for canvas-rendered content controls.
 *
 * The accessibility mirror must remain invisible and pointer-inert. This module
 * therefore derives a separate, visible DOM layer from the same display-list
 * metadata. React and Vue mount the returned element above a canvas page and
 * keep their existing delegated `.layout-sdt-widget` / repeat handlers.
 *
 * Security: every DOCX-derived string is assigned through `textContent`,
 * `dataset`, or `setAttribute`; no HTML strings are parsed.
 */
import type { DisplayPage, DisplayPrimitive, InlineSdtWidgetAttrs, SdtAttrs } from './displayList';
import { type GeoRect } from './displayListGeometry';
export interface InteractiveOverlayLabels {
    /** Accessible name for a repeating-section add button. */
    addRepeatingItem?: string;
    /** Accessible name for a repeating-section remove button. */
    removeRepeatingItem?: string;
    /** Accessible name used when a control has no authored alias or tag. */
    control?: string;
}
export interface BuildInteractiveOverlayOptions {
    /** Document used to create elements (default: global `document`). */
    document?: Document;
    /** Host-localized button names. */
    labels?: InteractiveOverlayLabels;
}
/** Apply content-control focus to the interactive overlay boxes. */
export declare function applyInteractiveSdtFocus(root: ParentNode, focusedGroupIds: ReadonlySet<string>): void;
interface SdtExtent {
    attrs: SdtAttrs;
    rect: GeoRect;
}
interface WidgetExtent {
    attrs: InlineSdtWidgetAttrs;
    rect: GeoRect;
    primitives: DisplayPrimitive[];
}
interface ControlExtent {
    rect: GeoRect;
    occluded: boolean;
}
interface BoundaryExtent extends SdtExtent {
    kind: ReturnType<typeof blockWidgetKind>;
    widget?: ControlExtent;
    repeats: ControlExtent[];
}
interface InteractiveOverlayLayer {
    lowerLayer: boolean;
    boundaries: BoundaryExtent[];
    widgets: Array<WidgetExtent & {
        occluded: boolean;
    }>;
}
/**
 * Build the visible/pointer-active overlay for one page.
 *
 * Coordinates remain page-local CSS pixels. The caller applies the same zoom
 * transform as the canvas and mirror. Boundary boxes themselves never receive
 * pointer events; only real buttons are interactive/focusable.
 */
export declare function buildInteractiveOverlayPage(page: DisplayPage, options?: BuildInteractiveOverlayOptions): HTMLElement;
export declare function collectInteractiveOverlayLayers(page: DisplayPage): InteractiveOverlayLayer[];
export declare function hideOccludedControl(element: HTMLElement): void;
declare function blockWidgetKind(sdtType: string): 'checkbox' | 'dropdown' | 'date' | null;
/** Whether `buildInteractiveOverlayPage(page)` holds a control Tab stops at. */
export declare function interactiveOverlayHasTabStops(page: DisplayPage): boolean;
export {};
