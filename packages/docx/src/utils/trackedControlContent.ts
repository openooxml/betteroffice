import type { ParagraphContent } from '../types/document';

export function visitTrackedControlContent(
  node: ParagraphContent,
  visit: (node: ParagraphContent) => void,
  control = false,
  revision = false
): void {
  if (node.type === 'inlineSdt') control = true;
  if (
    node.type === 'insertion' || node.type === 'deletion' ||
    node.type === 'moveFrom' || node.type === 'moveTo'
  ) revision = true;
  if (control && revision) visit(node);
  let children: readonly ParagraphContent[];
  if (
    node.type === 'inlineSdt' || node.type === 'insertion' || node.type === 'deletion' ||
    node.type === 'moveFrom' || node.type === 'moveTo'
  ) children = node.content;
  else if (node.type === 'hyperlink') children = node.structuredChildren ?? node.children;
  else if (node.type === 'simpleField') children = node.structuredResult?.inline ?? node.content;
  else if (node.type === 'complexField') {
    for (const child of node.structuredCode?.inline ?? node.fieldCode) {
      if (child.type === 'run' && !(control && revision)) continue;
      visitTrackedControlContent(child, visit, control, revision);
    }
    children = node.structuredResult?.inline ?? node.fieldResult;
  } else return;
  for (const child of children) {
    if (child.type === 'run' && !(control && revision)) continue;
    visitTrackedControlContent(child, visit, control, revision);
  }
}

export function hasTrackedControlContent(node: ParagraphContent): boolean {
  let found = false;
  visitTrackedControlContent(node, () => { found = true; });
  return found;
}
