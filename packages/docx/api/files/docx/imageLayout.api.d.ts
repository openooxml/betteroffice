// @betteroffice/docx declaration docx/imageLayout.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { WrapType } from './wrapTypes';
export type AnchorWrapType = Exclude<WrapType, 'inline'>;
export type ImageLayoutTarget = AnchorWrapType | 'squareLeft' | 'squareRight' | 'inline';
export type ImageResizeHandle = 'nw' | 'ne' | 'se' | 'sw' | 'n' | 's' | 'e' | 'w';
/** Compute the preview dimensions for an image resize gesture. */
export declare function calculateResizedImageDimensions(handle: ImageResizeHandle, deltaX: number, deltaY: number, startWidth: number, startHeight: number, lockAspect: boolean): {
    width: number;
    height: number;
};
export interface SetImageWrapTypeOptions {
    initialPositionEmu?: {
        horizontalEmu: number;
        verticalEmu: number;
    };
}
export interface ImageLayoutPosition {
    horizontal?: {
        relativeTo?: string;
        posOffset?: number;
        align?: string;
    };
    vertical?: {
        relativeTo?: string;
        posOffset?: number;
        align?: string;
    };
}
export interface ImageLayoutCurrentAttrs {
    wrapType?: string | null;
    cssFloat?: string | null;
    position?: ImageLayoutPosition | null;
}
export interface ImageLayoutAttrs extends Record<string, unknown> {
    wrapType: WrapType;
    displayMode: 'inline' | 'float' | 'block';
    cssFloat: 'left' | 'right' | 'none';
    wrapText?: 'bothSides' | 'left' | 'right' | 'largest';
    position?: ImageLayoutPosition;
    distTop?: number;
    distBottom?: number;
    distLeft?: number;
    distRight?: number;
}
/** Resolve the authored image attributes for one Word-style wrap choice. */
export declare function resolveImageLayoutAttrs(target: ImageLayoutTarget, current: ImageLayoutCurrentAttrs, opts?: SetImageWrapTypeOptions): ImageLayoutAttrs;
/** Map current OOXML image attrs to the shared menu choice vocabulary. */
export declare function deriveLayoutChoice(wrapType: WrapType, cssFloat?: string | null): ImageLayoutTarget | null;
export type ImageLayoutIconHint = 'inline' | 'squareLeft' | 'squareRight' | 'behind' | 'inFront';
export interface ImageLayoutOptionDef {
    choice: ImageLayoutTarget;
    i18nLabelKey: string;
    i18nDescKey: string;
    iconHint: ImageLayoutIconHint;
}
/** Word's five directional Wrap Text menu entries. */
export declare const IMAGE_LAYOUT_OPTIONS: readonly ImageLayoutOptionDef[];
export declare function isImageLayoutOptionEnabled(_option: ImageLayoutOptionDef, _currentWrapType: WrapType): boolean;
/** Map legacy toolbar wrap values onto the shared image command target. */
export declare function toolbarValueToLayoutTarget(value: string): ImageLayoutTarget | undefined;
