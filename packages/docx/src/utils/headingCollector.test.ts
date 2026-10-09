import { expect, test } from 'bun:test';
import type { EditorTreeNode } from '../types/editorTree';
import { collectHeadings } from './headingCollector';

function node(name: string, attrs: Record<string, unknown>, children: EditorTreeNode[] = [], text?: string) {
  return {
    type: { name, create: () => undefined as never },
    attrs,
    marks: [],
    nodeSize: 1,
    content: { size: children.length },
    isText: name === 'text',
    isTextblock: name === 'paragraph',
    isInline: name === 'text',
    text,
    forEach: (callback: (child: EditorTreeNode, offset: number, index: number) => void) =>
      children.forEach((child, index) => callback(child, index, index)),
    descendants: (callback: (child: EditorTreeNode, pos: number) => boolean | void) =>
      children.forEach((child, index) => callback(child, index * 10)),
  } as EditorTreeNode;
}

const paragraph = (attrs: Record<string, unknown>, text: string) =>
  node('paragraph', attrs, [node('text', {}, [], text)]);

test('the legacy collector reads outline levels and HeadingN style ids without wasm', () => {
  const doc = node('doc', {}, [
    paragraph({ outlineLevel: 1 }, 'Direct'),
    paragraph({ styleId: 'Heading2' }, 'Built-in'),
    paragraph({ outlineLevel: 9, styleId: 'Heading1' }, 'Body level'),
    paragraph({ styleId: 'Title' }, 'Plain'),
    paragraph({ outlineLevel: 0 }, '   '),
  ]);
  expect(collectHeadings(doc)).toEqual([
    { text: 'Direct', level: 1, pmPos: 0 },
    { text: 'Built-in', level: 1, pmPos: 10 },
  ]);
});
