import type { PptxPresencePeer } from '@betteroffice/pptx';

export const MAX_PRESENCE_CHIPS = 12;
export const MAX_THUMBNAIL_PRESENCE_DOTS = 4;
export const MAX_REMOTE_SHAPE_OUTLINES = 24;

export interface BoundedPresence<T> {
  visible: T[];
  overflow: number;
}

export interface ShapePresenceGroup {
  shapeId: string;
  peer: PptxPresencePeer;
  count: number;
}

interface IdLookup {
  has(id: string): boolean;
}

export function limitPresence<T>(
  values: readonly T[],
  limit = MAX_PRESENCE_CHIPS
): BoundedPresence<T> {
  return {
    visible: values.slice(0, limit),
    overflow: Math.max(0, values.length - limit),
  };
}

export function groupPresenceBySlide(
  peers: readonly PptxPresencePeer[],
  slideIds: IdLookup,
  limit = MAX_THUMBNAIL_PRESENCE_DOTS
): Map<string, BoundedPresence<PptxPresencePeer>> {
  const grouped = new Map<string, BoundedPresence<PptxPresencePeer>>();
  for (const peer of peers) {
    const slideId = peer.state.cursor?.slideId;
    if (!slideId || !slideIds.has(slideId)) continue;
    const group = grouped.get(slideId) ?? { visible: [], overflow: 0 };
    if (group.visible.length < limit) group.visible.push(peer);
    else group.overflow += 1;
    grouped.set(slideId, group);
  }
  return grouped;
}

export function groupShapePresence(
  peers: readonly PptxPresencePeer[],
  slideId: string,
  shapeIds: IdLookup,
  limit = MAX_REMOTE_SHAPE_OUTLINES
): BoundedPresence<ShapePresenceGroup> {
  const groups = new Map<string, ShapePresenceGroup>();
  let overflow = 0;
  for (const peer of peers) {
    const cursor = peer.state.cursor;
    if (cursor?.slideId !== slideId || !cursor.shapeId || !shapeIds.has(cursor.shapeId)) {
      continue;
    }
    const existing = groups.get(cursor.shapeId);
    if (existing) {
      existing.count += 1;
      if (peer.cursorMovedAt > existing.peer.cursorMovedAt) existing.peer = peer;
    } else if (groups.size < limit) {
      groups.set(cursor.shapeId, { shapeId: cursor.shapeId, peer, count: 1 });
    } else {
      overflow += 1;
    }
  }
  return { visible: [...groups.values()], overflow };
}
