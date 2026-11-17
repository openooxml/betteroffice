// @betteroffice/docx declaration yrs/yrsSessionFacade.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { YrsDocxHost, YrsSession } from './index';
import type { EditSession } from './wasm/index';
import type { DocumentBody } from '../types/document';
/** SHA-256 digests of the byte copies {@link prepareDocxBytes} made, by copy. */
/** @internal */
export declare const preparedDigests: WeakMap<Uint8Array<ArrayBufferLike>, string>;
/** @internal */
export declare function randomClientId(): number;
/**
 * Decodes the host metadata a resident worker's `open` replied with, for the
 * package `source` it opened. @internal
 */
export declare function decodeDocxHostJson(json: string, source: Uint8Array): YrsDocxHost;
export declare function seedDocxCommentMetadata(session: EditSession, comments: DocumentBody['comments']): void;
export declare function forgetMaterializedSource(session: EditSession): void;
/** @internal */
export declare function wrapSession(session: EditSession, clientId: number, opened?: {
    source: Uint8Array;
    host: YrsDocxHost;
}, loadPngExporter?: () => Promise<typeof import('./renderPng')>): YrsSession;
/** Borrows an opened edit session for saving; never destroy the facade. @internal */
export declare function wrapOpenedEditSession(session: EditSession, clientId: number, source: Uint8Array, hostJson: string): {
    session: YrsSession;
    host: YrsDocxHost;
};
