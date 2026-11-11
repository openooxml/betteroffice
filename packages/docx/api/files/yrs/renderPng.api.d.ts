// @betteroffice/docx declaration yrs/renderPng.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { type DocxPngOptions, type DocxPngTarget } from '../layout/render/png';
import { type YrsSession } from './index';
/**
 * Lays a replica of `session` out once per version with the editor's chain and
 * draws its pages. The replica opens from the session's own state, so it holds
 * the model the session has, and the session's retained layout stays untouched.
 */
export declare function createPngExporter(session: YrsSession): {
    render(target: DocxPngTarget, options?: DocxPngOptions): Promise<Uint8Array>;
    dispose(): void;
};
