/**
 * Structural mark contract consumed by the legacy tree-to-layout adapter.
 * @public
 */
export interface EditorTreeMark {
  readonly type: { readonly name: string };
  readonly attrs: Record<string, any>;
}

/**
 * Structural node contract consumed by renderer-side tree walkers.
 *
 * The contract deliberately describes only the immutable tree operations the
 * renderer needs, so core does not depend on any editor implementation.
 * @public
 */
export interface EditorTreeNode {
  readonly type: {
    readonly name: string;
    create(...args: any[]): EditorTreeNode;
  };
  readonly attrs: Record<string, any>;
  readonly marks: readonly EditorTreeMark[];
  readonly nodeSize: number;
  readonly content: { readonly size: number };
  readonly isText: boolean;
  readonly isTextblock: boolean;
  readonly isInline: boolean;
  readonly text?: string | null;
  forEach(callback: (node: EditorTreeNode, offset: number, index: number) => void): void;
  descendants(callback: (node: EditorTreeNode, pos: number) => boolean | void): void;
}
