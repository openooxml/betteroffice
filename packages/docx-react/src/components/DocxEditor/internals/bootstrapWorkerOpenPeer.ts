import { PeerMetadataError, peerMetadataTags, type YrsDocxHost, type YrsSession } from '@betteroffice/docx/yrs';

export function bootstrapWorkerOpenPeer(
  session: YrsSession,
  snapshot: { state: Uint8Array; metadata?: Uint8Array; metadataReason?: string },
  source: Uint8Array,
  host: YrsDocxHost,
): boolean {
  let reason = snapshot.metadataReason ?? 'missing-capability: Worker omitted peer metadata';
  if (snapshot.metadata !== undefined) {
    try {
      session.bootstrapPeer(snapshot.state, snapshot.metadata, source, host);
      return true;
    } catch (error) {
      if (!(error instanceof PeerMetadataError)) throw error;
      reason = `${error.code}: ${error.message}`;
    }
  }
  const tags = peerMetadataTags(snapshot.metadata);
  console.warn(`[yrs] peer hydration compatibility mode: ${reason}; expected tag ${tags.expected}; received tag ${tags.received}`);
  session.openDocx(source, false);
  return false;
}
