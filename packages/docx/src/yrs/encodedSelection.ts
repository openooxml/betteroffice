import type { CollaborationCursor } from '../collaboration/types';

/** A session's `encoded_selection()` JSON as binary sticky positions. */
export function decodeEncodedSelection(json: string): CollaborationCursor | null {
  const encoded = JSON.parse(json) as { story: string; anchor: number[]; head: number[] } | null;
  return encoded
    ? {
        story: encoded.story,
        anchor: Uint8Array.from(encoded.anchor),
        head: Uint8Array.from(encoded.head),
      }
    : null;
}
