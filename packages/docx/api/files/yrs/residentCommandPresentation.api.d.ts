// @betteroffice/docx declaration yrs/residentCommandPresentation.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { type TransferResult } from '../../../../shared/office-session/protocol';
import type { DocxCommandEchoContext, DocxCommandFrame, DocxCommandPageIdentity, DocxCommandStagedPresentation } from './residentCommandProtocol';
import type { DocxCommandSelectionLayout } from './residentCommandSelection';
import type { DisplayPage } from '../layout/render/displayList';
import type { RetainedFrame } from '../layout/render/frameDelta';
export declare function closeCommandBitmap(bitmap: ImageBitmap): void;
export declare function discardCommandPresentation(frame: DocxCommandFrame | undefined): void;
/**
 * @internal The page an echo background rasters: everything but the echoed line, reduced to what can
 * paint into the damage rect, so the background costs the rect rather than the page.
 */
export declare function echoBackgroundPage(retained: RetainedFrame['pages'][number], dependency: DocxCommandEchoContext['damage'][number]): DisplayPage;
/** @internal Worker-owned raster inputs; output bitmaps transfer to main. */
export interface StageCommandPresentationInput {
    layout: DocxCommandSelectionLayout;
    pages: readonly DocxCommandPageIdentity[];
    pixelRatio: number;
    echoContext?: DocxCommandEchoContext;
    rasterizePage(page: DocxCommandPageIdentity): Promise<ImageBitmap>;
    rasterizeBackground(dependency: DocxCommandEchoContext['damage'][number]): Promise<ImageBitmap>;
}
/** @internal */
export declare function stageCommandPresentation(input: StageCommandPresentationInput): Promise<TransferResult<DocxCommandStagedPresentation>>;
