// @betteroffice/docx declaration yrs/residentCommandConnection.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DocxCommandClient, DocxCommandConnection, DocxCommandInitReply, DocxCommandInitRequest } from './residentCommandProtocol';
export declare function createDocxCommandConnection(client: DocxCommandClient, init: DocxCommandInitRequest | DocxCommandInitReply): Promise<DocxCommandConnection>;
