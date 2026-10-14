import { PresentationPeerError } from '../wasm/loader';
import { PptxPeerHydrationError } from './peerHydrationError';

export function peerHydrationError(cause: unknown): unknown {
  if (cause instanceof PptxPeerHydrationError) return cause;
  if (cause instanceof PresentationPeerError) {
    return new PptxPeerHydrationError(cause.code, cause.message, cause);
  }
  if (cause instanceof Error && 'refusal' in cause) {
    const refusal = cause.refusal;
    if (refusal !== null && typeof refusal === 'object' && 'code' in refusal && typeof refusal.code === 'string') {
      return new PptxPeerHydrationError(refusal.code, cause.message, cause);
    }
  }
  return cause;
}
