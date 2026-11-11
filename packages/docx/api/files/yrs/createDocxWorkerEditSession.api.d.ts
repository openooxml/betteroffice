// @betteroffice/docx declaration yrs/createDocxWorkerEditSession.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DocxWorkerEditSession, DocxWorkerEditSessionOptions } from './residentCommandProtocol';
export declare function createDocxWorkerEditSession<O, X = never>(options: DocxWorkerEditSessionOptions<O, X>): DocxWorkerEditSession<X>;
