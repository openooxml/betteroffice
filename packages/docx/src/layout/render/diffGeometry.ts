// structured diff between two page-geometry snapshots (or a snapshot and a
// raw DisplayPage). nodes are matched by identity — region + kind + blockId +
// content discriminator (text / deco / edge role / relId) — deliberately EXCLUDING
// geometry and doc range, so a moved node reports a rect delta and a
// re-anchored node reports a doc-range mismatch instead of both degrading
// into a missing/extra pair. sub-pixel default tolerance.

import type { DisplayPage } from './displayList';
import {
  snapshotFromDisplayPage,
  type GeometryNode,
  type GeoRect,
  type PageGeometrySnapshot,
} from './harvestGeometry';

export const DEFAULT_TOLERANCE_PX = 0.5;

export type GeometryMismatch =
  | {
      type: 'page-size-mismatch';
      a: { width: number; height: number };
      b: { width: number; height: number };
    }
  | { type: 'missing-node'; side: 'a' | 'b'; key: string; node: GeometryNode }
  | { type: 'rect-delta'; key: string; a: GeoRect; b: GeoRect; maxDeltaPx: number }
  | {
      type: 'doc-range-mismatch';
      key: string;
      a: { docStart?: number; docEnd?: number };
      b: { docStart?: number; docEnd?: number };
    }
  | {
      type: 'attr-mismatch';
      key: string;
      field: 'commentIds' | 'revision' | 'text';
      a: unknown;
      b: unknown;
    };

export interface GeometryDiffReport {
  mismatches: GeometryMismatch[];
  /** node pairs that matched by identity (they may still carry mismatches) */
  matchedCount: number;
  tolerancePx: number;
}

export interface DiffGeometryOptions {
  tolerancePx?: number;
}

/**
 * diff two geometry snapshots; either side may be a raw DisplayPage, which is
 * normalized via snapshotFromDisplayPage first. an empty `mismatches` array
 * means the two renderings agree within tolerance.
 */
export function diffGeometry(
  a: PageGeometrySnapshot | DisplayPage,
  b: PageGeometrySnapshot | DisplayPage,
  options: DiffGeometryOptions = {}
): GeometryDiffReport {
  const tolerancePx = options.tolerancePx ?? DEFAULT_TOLERANCE_PX;
  const snapA = normalize(a);
  const snapB = normalize(b);
  const mismatches: GeometryMismatch[] = [];

  if (
    Math.abs(snapA.width - snapB.width) > tolerancePx ||
    Math.abs(snapA.height - snapB.height) > tolerancePx
  ) {
    mismatches.push({
      type: 'page-size-mismatch',
      a: { width: snapA.width, height: snapA.height },
      b: { width: snapB.width, height: snapB.height },
    });
  }

  const groupsA = groupByIdentity(snapA.nodes);
  const groupsB = groupByIdentity(snapB.nodes);
  let matchedCount = 0;

  for (const [key, listA] of groupsA) {
    const listB = groupsB.get(key) ?? [];
    const paired = Math.min(listA.length, listB.length);
    for (let i = 0; i < paired; i++) {
      matchedCount++;
      compareNode(indexedKey(key, i, paired), listA[i], listB[i], tolerancePx, mismatches);
    }
    for (let i = paired; i < listA.length; i++) {
      mismatches.push({
        type: 'missing-node',
        side: 'b',
        key: indexedKey(key, i, listA.length),
        node: listA[i],
      });
    }
    for (let i = paired; i < listB.length; i++) {
      mismatches.push({
        type: 'missing-node',
        side: 'a',
        key: indexedKey(key, i, listB.length),
        node: listB[i],
      });
    }
  }
  for (const [key, listB] of groupsB) {
    if (groupsA.has(key)) continue;
    for (let i = 0; i < listB.length; i++) {
      mismatches.push({
        type: 'missing-node',
        side: 'a',
        key: indexedKey(key, i, listB.length),
        node: listB[i],
      });
    }
  }

  return { mismatches, matchedCount, tolerancePx };
}

function normalize(input: PageGeometrySnapshot | DisplayPage): PageGeometrySnapshot {
  return 'primitives' in input ? snapshotFromDisplayPage(input) : input;
}

/**
 * identity key: what a node IS, independent of where it sits or what PM range
 * it maps to. region-scoped — a header text run never matches a body (or
 * footer) text run: regions map to different PM docs, so cross-region
 * "matches" would compare unrelated doc ranges and blockIds.
 */
export function identityKeyOf(node: GeometryNode): string {
  const discriminator =
    node.kind === 'text'
      ? node.text
      : node.kind === 'decoration'
        ? node.deco
        : node.kind === 'edge'
          ? node.role
          : node.kind === 'image'
            ? node.relId
            : '';
  return `${node.region ?? 'body'}|${node.kind}|block:${node.blockId ?? '-'}|${discriminator ?? ''}`;
}

function indexedKey(key: string, index: number, count: number): string {
  return count > 1 ? `${key}#${index}` : key;
}

function groupByIdentity(nodes: GeometryNode[]): Map<string, GeometryNode[]> {
  const groups = new Map<string, GeometryNode[]>();
  for (const node of nodes) {
    const key = identityKeyOf(node);
    const list = groups.get(key);
    if (list) list.push(node);
    else groups.set(key, [node]);
  }
  return groups;
}

function compareNode(
  key: string,
  a: GeometryNode,
  b: GeometryNode,
  tolerancePx: number,
  out: GeometryMismatch[]
): void {
  const maxDeltaPx = Math.max(
    Math.abs(a.rect.x - b.rect.x),
    Math.abs(a.rect.y - b.rect.y),
    Math.abs(a.rect.w - b.rect.w),
    Math.abs(a.rect.h - b.rect.h)
  );
  if (maxDeltaPx > tolerancePx) {
    out.push({ type: 'rect-delta', key, a: a.rect, b: b.rect, maxDeltaPx });
  }

  if (a.docStart !== b.docStart || a.docEnd !== b.docEnd) {
    out.push({
      type: 'doc-range-mismatch',
      key,
      a: { docStart: a.docStart, docEnd: a.docEnd },
      b: { docStart: b.docStart, docEnd: b.docEnd },
    });
  }

  const commentsA = a.commentIds ?? [];
  const commentsB = b.commentIds ?? [];
  if (commentsA.length !== commentsB.length || commentsA.some((id, i) => id !== commentsB[i])) {
    out.push({ type: 'attr-mismatch', key, field: 'commentIds', a: a.commentIds, b: b.commentIds });
  }

  const revA = a.revision;
  const revB = b.revision;
  if (
    (revA === undefined) !== (revB === undefined) ||
    (revA &&
      revB &&
      (revA.author !== revB.author ||
        revA.date !== revB.date ||
        revA.revisionId !== revB.revisionId ||
        revA.kind !== revB.kind))
  ) {
    out.push({ type: 'attr-mismatch', key, field: 'revision', a: revA, b: revB });
  }
}
