// @betteroffice/docx declaration layout/render/displayListHyperlinks.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DisplayList } from './displayList';
import type { DisplayListTableRegion } from './displayListTables';
import { type GeoRect } from './displayListGeometry';
export interface DisplayListHyperlinkHit {
    pageIndex: number;
    region: 'body' | 'header' | 'footer';
    rId?: string;
    href: string;
    tooltip?: string;
    displayText: string;
    rect: GeoRect;
}
export declare function findDisplayListHyperlinkAtPoint(list: DisplayList, pageIndex: number, x: number, y: number, region?: DisplayListTableRegion): DisplayListHyperlinkHit | null;
