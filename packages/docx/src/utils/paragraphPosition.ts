/**
 * Minimal document-tree node contract used by paragraph position helpers.
 * It is intentionally structural so callers do not need an editor runtime
 * dependency merely to walk a document tree.
 * @public
 */
export interface ParagraphPositionNode {
  readonly attrs: Record<string, unknown>;
  readonly isTextblock: boolean;
  readonly nodeSize: number;
  descendants(
    callback: (node: ParagraphPositionNode, pos: number) => boolean | void
  ): void;
}

/**
 * Position immediately before the first textblock whose `paraId` attribute
 * equals `paraId`.
 * @public
 */
export function findStartPosForParaId(
  doc: ParagraphPositionNode,
  paraId: string
): number | null {
  if (!paraId || !paraId.trim()) return null;
  let found: number | null = null;
  doc.descendants((node, pos) => {
    if (found !== null) return false;
    if (node.attrs?.paraId === paraId && node.isTextblock) {
      found = pos;
      return false;
    }
    return true;
  });
  return found;
}

/**
 * Position range for the textblock whose `paraId` attribute equals `paraId`.
 * @public
 */
export function findParagraphByParaId(
  doc: ParagraphPositionNode,
  paraId: string
): { node: ParagraphPositionNode; from: number; to: number } | null {
  if (!paraId || !paraId.trim()) return null;
  let result: { node: ParagraphPositionNode; from: number; to: number } | null = null;
  doc.descendants((node, pos) => {
    if (result !== null) return false;
    if (node.isTextblock && node.attrs?.paraId === paraId) {
      result = { node, from: pos, to: pos + node.nodeSize };
      return false;
    }
    return true;
  });
  return result;
}
