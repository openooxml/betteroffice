// @betteroffice/docx declaration yrs/peerMetadata.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** @internal */
export declare function peerMetadataTags(metadata?: Uint8Array): {
    expected: string;
    received: string;
};
export declare class PeerMetadataError extends Error {
    readonly code: string;
    constructor(code: string, message: string);
}
export declare function checkPeerMetadataHeader(metadata: Uint8Array): void;
