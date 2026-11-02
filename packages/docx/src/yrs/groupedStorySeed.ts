import type { Document, Shape } from '../types/document';

type ShapeOwner = { type: 'shape'; shape: Shape };

function visit(value: unknown, group: (owner: ShapeOwner) => void): void {
  if (value instanceof Map) {
    for (const child of value.values()) visit(child, group);
  } else if (Array.isArray(value)) {
    for (const child of value) visit(child, group);
  } else if (value && typeof value === 'object' && !ArrayBuffer.isView(value) && !(value instanceof ArrayBuffer)) {
    const object = value as Record<string, unknown>;
    const shape = object.type === 'shape' ? object.shape as Shape | undefined : undefined;
    if (shape?.children?.length) group(object as ShapeOwner);
    else for (const child of Object.values(object)) visit(child, group);
  }
}

function sourceJson(key: string, value: unknown): unknown {
  return key === 'sourceOrdinal' ? undefined : value;
}

function placeholderJson(key: string, value: unknown): unknown {
  if (key === 'textBody' && value && typeof value === 'object') {
    const body = value as Record<string, unknown>;
    if (Array.isArray(body.content)) {
      return { ...body, content: body.content.map(() => ({ type: 'paragraph', formatting: {}, content: [] })) };
    }
  }
  return sourceJson(key, value);
}

function stable(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stable).join(',')}]`;
  if (value && typeof value === 'object') {
    const object = value as Record<string, unknown>;
    return `{${Object.keys(object).sort().map(key => `${JSON.stringify(key)}:${stable(object[key])}`).join(',')}}`;
  }
  return JSON.stringify(value);
}

function signature(shape: Shape, placeholders: boolean): string {
  return stable(JSON.parse(JSON.stringify(shape, placeholders ? placeholderJson : sourceJson)));
}

export function recoverGroupedStorySeeds(document: Document, source: () => Readonly<Document> | null): () => void {
  let originals: Map<string, Shape | null> | undefined;
  const restored: Array<() => void> = [];
  visit(document.package, owner => {
    const key = signature(owner.shape, false);
    if (key !== signature(owner.shape, true)) return;
    if (!originals) {
      originals = new Map();
      const retained = source();
      if (retained) visit(retained.package, original => {
        const key = signature(original.shape, true);
        const previous = originals!.get(key);
        if (previous === undefined) originals!.set(key, original.shape);
        else if (JSON.stringify(previous, sourceJson) !== JSON.stringify(original.shape, sourceJson)) originals!.set(key, null);
      });
    }
    const original = originals.get(key);
    if (!original) return;
    const previous = owner.shape;
    owner.shape = JSON.parse(JSON.stringify(original, sourceJson)) as Shape;
    restored.push(() => { owner.shape = previous; });
  });
  return () => { for (const restore of restored) restore(); };
}
