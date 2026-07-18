/**
 * Repeating-section (w15:repeatingSection) support — add and remove repeated
 * items, the way Word's "+" affordance does. A repeating section is a block
 * content control whose `w:sdtPr` carries `<w15:repeatingSection>`; its direct
 * children are item controls each carrying `<w15:repeatingSectionItem>`.
 *
 * Adding clones an existing item (with a fresh, unique `w:id`) and inserts it
 * after; removing drops one item but keeps at least one. The w15 elements ride
 * in the captured raw `w:sdtPr` (they're unmodeled), so we detect and patch the
 * raw string rather than re-serializing.
 */

import type { Document, BlockContent, BlockSdt, SdtProperties } from '../types/document';
import {
  ContentControlNotFoundError,
  applyToFirst,
  rebuild,
  type ContentControlFilter,
  type ControlOp,
} from './contentControls';

/** Raw `w:sdtPr` declares a repeating section (the container). */
export function rawIsRepeatingSection(raw: string | undefined): boolean {
  // `repeatingSection` not followed by "Item" (the item element starts the same).
  return /<w15:repeatingSection[\s/>]/.test(raw ?? '');
}

/** Raw `w:sdtPr` declares a repeating-section item. */
export function rawIsRepeatingSectionItem(raw: string | undefined): boolean {
  return /<w15:repeatingSectionItem[\s/>]/.test(raw ?? '');
}

/** The control's raw `w:sdtPr` declares it a repeating section (the container). */
export function isRepeatingSection(props: SdtProperties): boolean {
  return rawIsRepeatingSection(props.rawPropertiesXml);
}

/** The control's raw `w:sdtPr` declares it a repeating-section item. */
export function isRepeatingSectionItem(props: SdtProperties): boolean {
  return rawIsRepeatingSectionItem(props.rawPropertiesXml);
}

/** Raised when an operation targets something that isn't a repeating section/item. */
export class RepeatingSectionError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'RepeatingSectionError';
  }
}

type TraversableControlNode = {
  type?: string;
  properties?: SdtProperties;
  [key: string]: unknown;
};

/** Walk every modeled child container, including inline controls and tables. */
function walkControlTree(value: unknown, visit: (control: TraversableControlNode) => void): void {
  if (Array.isArray(value)) {
    for (const child of value) walkControlTree(child, visit);
    return;
  }
  if (!value || typeof value !== 'object') return;
  const node = value as TraversableControlNode;
  if ((node.type === 'blockSdt' || node.type === 'inlineSdt') && node.properties) visit(node);
  for (const [key, child] of Object.entries(node)) {
    if (key === 'properties' || key === 'rawPropertiesXml' || key === 'rawEndPropertiesXml')
      continue;
    if (Array.isArray(child) || (child && typeof child === 'object')) {
      walkControlTree(child, visit);
    }
  }
}

/** Highest `w:id` used by any body/HF control, for minting a fresh one. */
function maxControlId(doc: Document): number {
  let max = 0;
  const inspect = (control: TraversableControlNode): void => {
    if (typeof control.properties?.id === 'number') max = Math.max(max, control.properties.id);
  };
  walkControlTree(doc.package.document.content, inspect);
  for (const part of doc.package.headers?.values() ?? []) walkControlTree(part.content, inspect);
  for (const part of doc.package.footers?.values() ?? []) walkControlTree(part.content, inspect);
  return max;
}

/** Patch (or leave) the `<w:id w:val="…">` inside a raw `w:sdtPr` string. */
export function patchRawId(raw: string | undefined, id: number): string | undefined {
  return setRawId(raw, id);
}
function setRawId(raw: string | undefined, id: number): string | undefined {
  if (!raw) return raw;
  if (/<w:id\b/.test(raw)) return raw.replace(/(<w:id\b[^>]*\bw:val=")[^"]*(")/, `$1${id}$2`);
  if (/<w:sdtPr\b[^>]*\/>/.test(raw)) {
    return raw.replace(/<w:sdtPr\b([^>]*)\/>/, `<w:sdtPr$1><w:id w:val="${id}"/></w:sdtPr>`);
  }
  return raw.replace(/(<w:sdtPr\b[^>]*[^/]>)/, `$1<w:id w:val="${id}"/>`);
}

/** Deep-clone an item control, assigning fresh unique ids to it and any nested controls. */
function cloneItem(item: BlockSdt, nextId: () => number): BlockSdt {
  const clone = structuredClone(item) as BlockSdt;
  walkControlTree(clone, (sdt) => {
    const id = nextId();
    const properties = sdt.properties!;
    sdt.properties = {
      ...properties,
      id,
      rawPropertiesXml: setRawId(properties.rawPropertiesXml, id),
    };
  });
  return clone;
}

/** Find the section matching `filter`; throws if missing or not a repeating section. */
function locateSection(doc: Document, filter: ContentControlFilter): BlockSdt {
  let found: BlockSdt | undefined;
  const walk = (blocks: BlockContent[]): void => {
    for (const b of blocks) {
      if (found) return;
      if (b.type === 'blockSdt') {
        const p = b.properties;
        const ok =
          (filter.tag === undefined || p.tag === filter.tag) &&
          (filter.alias === undefined || p.alias === filter.alias) &&
          (filter.id === undefined || p.id === filter.id);
        if (ok) found = b;
        else walk(b.content);
      } else if (b.type === 'table') {
        for (const row of b.rows) for (const cell of row.cells) walk(cell.content);
      }
    }
  };
  walk(doc.package.document.content);
  if (!found) throw new ContentControlNotFoundError(filter);
  if (!isRepeatingSection(found.properties)) {
    throw new RepeatingSectionError(
      `Control ${filter.tag ?? filter.alias ?? filter.id} is not a repeating section.`
    );
  }
  return found;
}

/** Item children of a repeating section, with their index in `section.content`. */
function itemIndices(section: BlockSdt): number[] {
  const out: number[] = [];
  section.content.forEach((b, i) => {
    if (b.type === 'blockSdt' && isRepeatingSectionItem(b.properties)) out.push(i);
  });
  return out;
}

/**
 * Add a new repeating-section item, cloned from an existing one and inserted
 * after it. `afterIndex` is the item ordinal to clone/insert after (default:
 * the last item). Returns a new {@link Document}.
 */
export function addRepeatingSectionItem(
  doc: Document,
  filter: ContentControlFilter,
  options: { afterIndex?: number } = {}
): Document {
  const section = locateSection(doc, filter);
  const items = itemIndices(section);
  if (items.length === 0) {
    throw new RepeatingSectionError('Repeating section has no item to clone.');
  }
  const ord = options.afterIndex ?? items.length - 1;
  const srcContentIdx = items[Math.max(0, Math.min(ord, items.length - 1))];
  const template = section.content[srcContentIdx] as BlockSdt;

  let nextId = maxControlId(doc);
  const clone = cloneItem(template, () => ++nextId);

  const state = { done: false };
  const op: ControlOp = (s) => {
    const next = [...s.content];
    next.splice(srcContentIdx + 1, 0, clone);
    return [{ ...s, content: next }];
  };
  const content = applyToFirst(doc.package.document.content, filter, op, state);
  return rebuild(doc, content);
}

/**
 * Remove the repeating-section item at ordinal `index`. Keeps at least one item
 * (Word does not allow removing the last). Returns a new {@link Document}.
 */
export function removeRepeatingSectionItem(
  doc: Document,
  filter: ContentControlFilter,
  index: number
): Document {
  const section = locateSection(doc, filter);
  const items = itemIndices(section);
  if (items.length <= 1) {
    throw new RepeatingSectionError('Cannot remove the last item of a repeating section.');
  }
  if (index < 0 || index >= items.length) {
    throw new RepeatingSectionError(`Item index ${index} out of range (0..${items.length - 1}).`);
  }
  const removeContentIdx = items[index];
  const state = { done: false };
  const op: ControlOp = (s) => {
    const next = s.content.filter((_, i) => i !== removeContentIdx);
    return [{ ...s, content: next }];
  };
  const content = applyToFirst(doc.package.document.content, filter, op, state);
  return rebuild(doc, content);
}
