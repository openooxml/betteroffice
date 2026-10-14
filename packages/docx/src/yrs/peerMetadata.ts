const MAGIC = [66, 79, 80, 69, 69, 82, 0, 0];
const VERSION = 1;
const HEADER_LENGTH = 60;
const SHAPE = '99297429e552981d11c59676fbec625b869187d6222616525f083618550f250a';

/** @internal */
export function peerMetadataTags(metadata?: Uint8Array): { expected: string; received: string } {
  const expected = `v${VERSION}/${SHAPE}`;
  if (!metadata) return { expected, received: 'absent' };
  if (metadata.byteLength < 44) return { expected, received: `truncated (${metadata.byteLength} bytes)` };
  const version = new DataView(metadata.buffer, metadata.byteOffset, metadata.byteLength).getUint32(8, true);
  const shape = Array.from(metadata.subarray(12, 44), (byte) => byte.toString(16).padStart(2, '0')).join('');
  return { expected, received: `v${version}/${shape}` };
}

export class PeerMetadataError extends Error {
  constructor(readonly code: string, message: string) {
    super(message);
    this.name = 'PeerMetadataError';
  }
}

export function checkPeerMetadataHeader(metadata: Uint8Array): void {
  if (metadata.byteLength < HEADER_LENGTH) {
    throw new PeerMetadataError('truncated', 'Truncated peer metadata');
  }
  if (MAGIC.some((byte, at) => metadata[at] !== byte)) {
    throw new PeerMetadataError('bad-magic', 'Invalid peer metadata magic');
  }
  const view = new DataView(metadata.buffer, metadata.byteOffset, metadata.byteLength);
  const version = view.getUint32(8, true);
  if (version !== VERSION) {
    throw new PeerMetadataError('unsupported-version', `Unsupported peer metadata version ${version}`);
  }
  for (let at = 0; at < 32; at += 1) {
    if (metadata[12 + at] !== Number.parseInt(SHAPE.slice(at * 2, at * 2 + 2), 16)) {
      throw new PeerMetadataError('shape-mismatch', 'Peer metadata shape fingerprint mismatch');
    }
  }
  const length = BigInt(HEADER_LENGTH) + view.getBigUint64(44, true) + view.getBigUint64(52, true);
  if (length > BigInt(Number.MAX_SAFE_INTEGER)) {
    throw new PeerMetadataError('invalid-length', 'Invalid peer metadata section length');
  }
  if (length > BigInt(metadata.byteLength)) {
    throw new PeerMetadataError('truncated', 'Truncated peer metadata');
  }
  if (length !== BigInt(metadata.byteLength)) {
    throw new PeerMetadataError('invalid-length', 'Invalid peer metadata section length');
  }
}
