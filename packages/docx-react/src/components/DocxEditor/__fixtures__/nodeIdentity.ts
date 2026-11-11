const ATTRIBUTES = ['data-block-id', 'data-para-id', 'role', 'class', 'data-docx-unbuilt-page', 'data-page-index'];

/** A short description of `node` for an assertion message; a DOM node itself prints its whole object graph. */
export function describeNode(node: unknown): string {
  if (node === null || node === undefined) return String(node);
  if (typeof node !== 'object' || !('nodeType' in node)) return typeof node;
  const element = node as Element;
  if (element.nodeType !== 1) return `#${element.nodeName}${element.isConnected ? '' : ' (detached)'}`;
  const attributes = ATTRIBUTES.map((name) => element.getAttribute(name) === null ? '' : `[${name}="${element.getAttribute(name)}"]`).join('');
  return `${element.tagName}${attributes}${element.isConnected ? '' : ' (detached)'}`;
}

/** 'same' when `actual` is `expected`, else what `actual` is. */
export function sameNode(actual: unknown, expected: unknown): string {
  return actual === expected ? 'same' : `${describeNode(actual)}, expected ${describeNode(expected)}`;
}
