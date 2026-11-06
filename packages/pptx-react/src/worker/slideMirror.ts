import type { ShapeAltText, SlideDisplayList, SlidePrimitive, TextLanguage } from '@betteroffice/pptx';

export interface SlideMirrorSegment {
  text: string;
  lang: string | null;
}

export interface SlideMirrorParagraph {
  text: string;
  heading: boolean;
  lang: string | null;
  segments: SlideMirrorSegment[];
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
const spoken = (text: string) => text.replace(/[\s\u000b]+/g, ' ');

const midWord = (before: string, after: string) => {
  const unspaced = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Thai}]/u;
  const [tail] = /[\p{L}\p{N}\p{M}]$/u.exec(before) ?? [];
  const [head] = /^[\p{L}\p{N}]/u.exec(after) ?? [];
  return tail !== undefined && head !== undefined && !unspaced.test(tail) && !unspaced.test(head);
};
const letters = /\p{L}.*\p{L}/su;
const primaryTag = (tag: string) => tag.split('-')[0]!.toLowerCase();

/** Language tags land on arbitrary runs, so a change only counts between whole words of two letters or more. */
function wordAligned(parts: SlideMirrorSegment[]): SlideMirrorSegment[] {
  const aligned: SlideMirrorSegment[] = [];
  for (const part of parts) {
    let { text } = part;
    const previous = aligned[aligned.length - 1];
    if (previous && previous.lang !== part.lang) {
      if (midWord(previous.text, text)) {
        const [rest] = /^[\p{L}\p{N}\p{M}]+/u.exec(text)!;
        previous.text += rest;
        text = text.slice(rest.length);
      }
      if (!letters.test(text)) {
        previous.text += text;
        continue;
      }
      if (!letters.test(previous.text)) {
        previous.text += text;
        previous.lang = part.lang;
        continue;
      }
    } else if (previous) {
      previous.text += text;
      continue;
    }
    aligned.push({ text, lang: part.lang });
  }
  return aligned;
}

function mirrorParagraph(source: SlideMirrorSegment[], heading = false): SlideMirrorParagraph | null {
  const parts = wordAligned(source);
  const counts = new Map<string, number>();
  for (const part of parts) if (part.lang) counts.set(part.lang, (counts.get(part.lang) ?? 0) + part.text.replace(/\P{L}/gu, '').length);
  let lang: string | null = null;
  let most = -1;
  if (parts.every((part) => part.lang)) for (const [tag, count] of counts) if (count > most) { lang = tag; most = count; }
  const segments: SlideMirrorSegment[] = [];
  for (const part of parts) {
    let text = spoken(part.text);
    const previous = segments[segments.length - 1];
    if (!previous || previous.text.endsWith(' ')) text = text.trimStart();
    if (!text) continue;
    const language = part.lang && lang && primaryTag(part.lang) === primaryTag(lang) ? lang : part.lang;
    if (previous && previous.lang === language) previous.text += text;
    else segments.push({ text, lang: language });
  }
  const last = segments[segments.length - 1];
  if (last) {
    last.text = last.text.trimEnd();
    if (!last.text) segments.pop();
  }
  const text = segments.map((segment) => segment.text).join('');
  return text ? { text, heading, lang, segments } : null;
}

function languageParts(text: string, languages: TextLanguage[] = []): SlideMirrorSegment[] {
  const parts: SlideMirrorSegment[] = [];
  let offset = 0;
  for (const language of languages) {
    const start = Math.max(offset, Math.min(text.length, language.start));
    const end = Math.max(start, Math.min(text.length, language.end));
    if (start > offset) parts.push({ text: text.slice(offset, start), lang: null });
    const tag = language.lang.trim();
    if (end > start) parts.push({ text: text.slice(start, end), lang: tag && tag.toLowerCase() !== 'x-none' ? tag : null });
    offset = end;
  }
  if (offset < text.length) parts.push({ text: text.slice(offset), lang: null });
  return parts;
}

function altLabel(entry: ShapeAltText | undefined): string | null {
  return [entry?.title, entry?.description].filter(Boolean).join(': ') || null;
}

function paragraphs(primitive: SlidePrimitive): SlideMirrorParagraph[] {
  if (primitive.kind === 'textBox') return primitive.paragraphs.flatMap((paragraph) => {
    const value = mirrorParagraph(languageParts(paragraph.runs.map((run) => run.text).join(''), paragraph.languages));
    return value ? [value] : [];
  });
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
    own.paragraphs.push(...paragraphs(primitive));
    if (entry?.decorative || primitive.kind === 'textBox') continue;
    const described = altLabel(entry);
    const artwork = primitive.kind === 'image' || painted.get(index) === id;
    const summary = primitive.kind === 'shape' || primitive.kind === 'image' ? null : primitive.label ?? primitive.name;
    const fallback = entry?.group ? null : artwork ? primitive.name || picture : summary;
    own.label ??= described ?? fallback;
  }
  const ordered = [...items.values()].filter((value) => value.label !== null || value.paragraphs.length > 0);
  const title = list.titleShape ? items.get(list.titleShape) : undefined;
  if (title?.paragraphs.length) {
    const parts = title.paragraphs.flatMap((paragraph, index) => [
      ...(index ? [{ text: ' ', lang: null }] : []), ...paragraph.segments,
    ]);
    title.paragraphs = [mirrorParagraph(parts, true)!];
  }
  return [...ordered.filter((value) => !inherited(value.key)), ...ordered.filter((value) => inherited(value.key))];
}

/** The text of the slide's title placeholder, when it has any. */
export function slideTitle(list: SlideDisplayList): string | null {
  const id = list.titleShape;
  if (!id) return null;
  return list.primitives.filter((primitive) => primitive.shapeId === id).flatMap(paragraphs).map((paragraph) => paragraph.text).join(' ') || null;
}

/** How an announcement names a shape: its alternative text, its name, then its text. */
export function slideObjectName(list: SlideDisplayList, shapeId: string, fallback: string): string {
  const entry = list.altText?.[shapeId];
  const described = entry?.decorative ? null : altLabel(entry);
  if (described) return described;
  const start = entry?.primitive;
  const primitives = list.primitives.filter((primitive, index) => primitive.shapeId === shapeId || index === start);
  for (const primitive of primitives) if (primitive.kind !== 'textBox' && primitive.name) return primitive.name;
  return primitives.flatMap(paragraphs)[0]?.text ?? fallback;
}
