import type { ShapeAltText, SlideDisplayList, SlidePrimitive } from '@betteroffice/pptx';

export interface SlideMirrorParagraph {
  text: string;
  heading: boolean;
}

/** One shape as a screen reader meets it, in slide pixels. */
export interface SlideMirrorItem {
  key: string;
  x: number;
  y: number;
  w: number;
  h: number;
  label: string | null;
  paragraphs: SlideMirrorParagraph[];
}

const inherited = (shapeId: string) => shapeId.startsWith('master:') || shapeId.startsWith('layout:');
const spoken = (text: string) => text.replace(/[\s\u000b]+/g, ' ').trim();

function altLabel(entry: ShapeAltText | undefined): string | null {
  return [entry?.title, entry?.description].filter(Boolean).join(': ') || null;
}

function paragraphs(primitive: SlidePrimitive): string[] {
  if (primitive.kind === 'textBox') return primitive.paragraphs.map((paragraph) => spoken(paragraph.runs.map((run) => run.text).join(''))).filter(Boolean);
  return primitive.kind === 'table' ? primitive.primitives.flatMap(paragraphs) : [];
}

/**
 * The slide in reading order: its own shapes back to front, as PowerPoint reads
 * them, then what the layout and master paint on every slide.
 */
export function slideMirrorItems(list: SlideDisplayList, picture: string): SlideMirrorItem[] {
  const alt = list.altText ?? {};
  const items = new Map<string, SlideMirrorItem>();
  const item = (key: string, { x, y, w, h }: SlidePrimitive) => {
    let value = items.get(key);
    if (!value) items.set(key, value = { key, x, y, w, h, label: null, paragraphs: [] });
    return value;
  };
  const describeGroups = (primitive: SlidePrimitive, group: string | undefined) => {
    if (!group || items.has(group)) return;
    describeGroups(primitive, alt[group]?.group);
    if (!alt[group]?.decorative) item(group, primitive).label = altLabel(alt[group]);
  };
  const painted = new Map(Object.entries(alt).filter(([, entry]) => entry.primitive !== undefined)
    .map(([id, entry]) => [entry.primitive, id] as const));
  for (const [index, primitive] of list.primitives.entries()) {
    const id = primitive.shapeId ?? painted.get(index);
    if (!id) continue;
    const entry = alt[id];
    describeGroups(primitive, entry?.group);
    const own = item(id, primitive);
    own.paragraphs.push(...paragraphs(primitive).map((text) => ({ text, heading: false })));
    if (entry?.decorative || primitive.kind === 'textBox') continue;
    const described = altLabel(entry);
    const artwork = primitive.kind === 'image' || painted.get(index) === id;
    const summary = primitive.kind === 'shape' || primitive.kind === 'image' ? null : primitive.label ?? primitive.name;
    const fallback = entry?.group ? null : artwork ? primitive.name || picture : summary;
    own.label ??= described ?? fallback;
  }
  const ordered = [...items.values()].filter((value) => value.label !== null || value.paragraphs.length > 0);
  const title = list.titleShape ? items.get(list.titleShape) : undefined;
  if (title?.paragraphs.length) title.paragraphs = [{ text: title.paragraphs.map((paragraph) => paragraph.text).join(' '), heading: true }];
  return [...ordered.filter((value) => !inherited(value.key)), ...ordered.filter((value) => inherited(value.key))];
}

/** The text of the slide's title placeholder, when it has any. */
export function slideTitle(list: SlideDisplayList): string | null {
  const id = list.titleShape;
  if (!id) return null;
  return list.primitives.filter((primitive) => primitive.shapeId === id).flatMap(paragraphs).join(' ') || null;
}

/** How an announcement names a shape: its alternative text, its name, then its text. */
export function slideObjectName(list: SlideDisplayList, shapeId: string, fallback: string): string {
  const entry = list.altText?.[shapeId];
  const described = entry?.decorative ? null : altLabel(entry);
  if (described) return described;
  const start = entry?.primitive;
  const primitives = list.primitives.filter((primitive, index) => primitive.shapeId === shapeId || index === start);
  for (const primitive of primitives) if (primitive.kind !== 'textBox' && primitive.name) return primitive.name;
  return primitives.flatMap(paragraphs)[0] ?? fallback;
}
