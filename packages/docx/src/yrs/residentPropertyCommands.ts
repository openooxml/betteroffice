import type { DocxCommandEnvelope, DocxCommandIntent, DocxCommandApplyResult, ResidentDocxCommandEngine } from './residentCommandProtocol';
import { commandSelectionRange } from './residentCommandAnchors';
import { patchSection, readResidentProperties, resolvePropertyImage, sectionProperties, watermarkDocument } from './residentProperties';

export const PROPERTY_COMMANDS = ['setHyperlink', 'setObjectProperties', 'setSectionProperties', 'setNoteProperties', 'setWatermark'] as const;
const keys: Record<(typeof PROPERTY_COMMANDS)[number], string[]> = {
  setHyperlink: ['selection', 'hyperlink', 'expectVersion'],
  setObjectProperties: ['handle', 'patch', 'expectVersion'],
  setSectionProperties: ['section', 'patch', 'expectVersion'],
  setNoteProperties: ['section', 'footnotePr', 'endnotePr', 'expectVersion'],
  setWatermark: ['watermark', 'expectVersion'],
};

function record(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value) &&
    (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
}
const finite = (value: unknown): value is number => typeof value === 'number' && Number.isFinite(value);
const text = (value: unknown): value is string => typeof value === 'string' && !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(value);
const choice = (value: unknown, values: string[]) => text(value) && values.includes(value);

function note(value: unknown, endnote = false): boolean {
  return record(value) && Object.entries(value).every(([key, entry]) => {
    if (key === 'numStart') return finite(entry) && Number.isSafeInteger(entry) && entry > 0;
    if (key === 'position') return choice(entry, endnote ? ['sectEnd', 'docEnd'] : ['pageBottom', 'beneathText', 'sectEnd', 'docEnd']);
    if (key === 'numRestart') return choice(entry, ['continuous', 'eachSect', 'eachPage']);
    if (key === 'numFmt') return choice(entry, ['decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter', 'ordinal', 'cardinalText', 'ordinalText', 'chicago', 'none']);
    if (key === 'customNumberFormat') return text(entry);
    if (key === 'separators') return Array.isArray(entry) && entry.every(separator => record(separator) &&
      Object.entries(separator).every(([k, v]) => k === 'kind' ? choice(v, ['separator', 'continuationSeparator', 'continuationNotice'])
        : k === 'noteId' && finite(v) && Number.isSafeInteger(v)));
    return false;
  });
}

function sectionPatch(value: unknown): boolean {
  return record(value) && Object.entries(value).every(([key, entry]) => {
    if (['pageWidth', 'pageHeight', 'columnCount', 'footnoteColumns'].includes(key)) return finite(entry) && entry > 0 && (key.endsWith('Count') || key.endsWith('Columns') ? Number.isSafeInteger(entry) : true);
    if (['marginLeft', 'marginRight', 'marginTop', 'marginBottom', 'headerDistance', 'footerDistance', 'gutter', 'columnSpace'].includes(key)) return finite(entry) && entry >= 0;
    if (key === 'orientation') return choice(entry, ['portrait', 'landscape']);
    if (key === 'sectionStart') return choice(entry, ['nextPage', 'continuous', 'oddPage', 'evenPage', 'nextColumn']);
    if (key === 'verticalAlign') return choice(entry, ['top', 'center', 'both', 'bottom']);
    if (['equalWidth', 'separator', 'bidi', 'titlePg', 'evenAndOddHeaders'].includes(key)) return typeof entry === 'boolean';
    if (key === 'footnotePr' || key === 'endnotePr') return note(entry, key === 'endnotePr');
    if (key === 'columns') return Array.isArray(entry) && entry.every(column => record(column) && Object.entries(column).every(([k, v]) => ['width', 'space'].includes(k) && finite(v) && v >= 0));
    return false;
  });
}

function imagePatch(value: unknown): boolean {
  return record(value) && Object.entries(value).every(([key, entry]) => {
    if (['widthEmu', 'heightEmu'].includes(key)) return finite(entry) && entry > 0;
    if (['width', 'height', 'borderWidth', 'distTop', 'distBottom', 'distLeft', 'distRight'].includes(key)) return entry === null || finite(entry) && entry >= 0;
    if (['hOffsetEmu', 'vOffsetEmu', 'distTopEmu', 'distBottomEmu', 'distLeftEmu', 'distRightEmu'].includes(key)) return entry === null || finite(entry);
    if (key === 'wrap') return choice(entry, ['inline', 'square', 'tight', 'through', 'topAndBottom', 'behind', 'inFront']);
    if (['alt', 'title', 'borderColor', 'relativeFromHorizontal', 'relativeFromVertical'].includes(key)) return entry === null || text(entry);
    if (key === 'borderColorValue') return entry === null;
    if (key === 'borderStyle') return entry === null || choice(entry, ['solid', 'dashed', 'dotted', 'double']);
    if (key === 'position') return entry === null || record(entry) && Object.entries(entry).every(([axis, position]) =>
      ['horizontal', 'vertical'].includes(axis) && record(position) && Object.entries(position).every(([k, v]) =>
        k === 'posOffset' ? finite(v) : ['relativeTo', 'align'].includes(k) && text(v)));
    if (key === 'decorative') return typeof entry === 'boolean';
    return false;
  });
}

function watermark(value: unknown): boolean {
  if (value === null) return true;
  if (!record(value)) return false;
  if (value.kind === 'text') return text(value.text) && text(value.font) && text(value.color) &&
    typeof value.semitransparent === 'boolean' && choice(value.layout, ['diagonal', 'horizontal']) &&
    Object.entries(value).every(([key, entry]) => ['kind', 'text', 'font', 'color', 'semitransparent', 'layout'].includes(key) ||
      (key === 'fontSize' && finite(entry) && entry > 0) || (key === 'decorative' && typeof entry === 'boolean'));
  if (value.kind === 'picture') return (text(value.dataUrl) || text(value.mediaPath) || text(value.relId) || value.data instanceof Uint8Array && value.data.length > 0) &&
    finite(value.scale) && value.scale > 0 && typeof value.washout === 'boolean' &&
    Object.entries(value).every(([key, entry]) => ['kind', 'scale', 'washout'].includes(key) ||
      (['relId', 'contentType', 'dataUrl', 'mediaPath'].includes(key) && text(entry)) || (key === 'data' && entry instanceof Uint8Array) ||
      (['widthEmu', 'heightEmu'].includes(key) && finite(entry) && entry > 0) ||
      (key === 'decorative' && typeof entry === 'boolean'));
  return false;
}

export function propertyCommand(kind: string): kind is (typeof PROPERTY_COMMANDS)[number] {
  return Object.hasOwn(keys, kind);
}

export function validPropertyCommand(command: DocxCommandEnvelope): boolean {
  if (!propertyCommand(command.kind) || !record(command.payload)) return false;
  const payload = command.payload as Record<string, unknown>;
  const allowed = keys[command.kind];
  if (!text(payload.expectVersion) || Object.keys(payload).some(key => !allowed.includes(key))) return false;
  switch (command.kind) {
    case 'setHyperlink': return record(payload.selection) && (payload.hyperlink === null ||
      record(payload.hyperlink) && text(payload.hyperlink.href) && Object.entries(payload.hyperlink).every(([key, entry]) => ['href', 'tooltip', 'rId'].includes(key) && text(entry)));
    case 'setObjectProperties': return text(payload.handle) && imagePatch(payload.patch);
    case 'setSectionProperties': return text(payload.section) && sectionPatch(payload.patch);
    case 'setNoteProperties': return text(payload.section) && note(payload.footnotePr) && note(payload.endnotePr, true);
    case 'setWatermark': return watermark(payload.watermark);
  }
}

export function propertyCommandTarget(engine: ResidentDocxCommandEngine, command: DocxCommandIntent): string | null {
  switch (command.kind) {
    case 'setHyperlink': {
      const range = commandSelectionRange(engine, command.payload.selection);
      return range && (range.start.paraId !== range.end.paraId || range.start.offset !== range.end.offset) ? null : 'invalid-range';
    }
    case 'setObjectProperties': return resolvePropertyImage(engine, command.payload.handle) ? null : 'stale-object';
    case 'setSectionProperties': case 'setNoteProperties': return sectionProperties(engine.propertyDocument(), command.payload.section) ? null : 'unknown-section';
    default: return null;
  }
}

export function applyPropertyCommand(engine: ResidentDocxCommandEngine, command: DocxCommandIntent): Pick<DocxCommandApplyResult, 'status'> {
  const before = engine.proposalEngine.version();
  engine.beginUndoCapture();
  engine.addUndoBoundary();
  switch (command.kind) {
    case 'setHyperlink': engine.setHyperlink(commandSelectionRange(engine, command.payload.selection)!, command.payload.hyperlink); break;
    case 'setObjectProperties': {
      const at = resolvePropertyImage(engine, command.payload.handle)!;
      const read = readResidentProperties(engine, { kind: 'object', at });
      if (read?.kind !== 'object') throw new Error('Image handle no longer resolves');
      const patch = command.payload.patch;
      if (Object.entries(patch).every(([key, value]) => value === null ? read.properties[key] === undefined
        : JSON.stringify(read.properties[key]) === JSON.stringify(value))) break;
      const other = { ...patch, ...(patch.widthEmu === undefined ? {} : { width: Number(patch.widthEmu) / 9525 }),
        ...(patch.heightEmu === undefined ? {} : { height: Number(patch.heightEmu) / 9525 }),
        ...(patch.wrap === undefined ? {} : { wrapType: patch.wrap }) };
      engine.setImageGeometryAt(at, { widthEmu: Number(patch.widthEmu ?? Number(patch.width ?? read.properties.width ?? 0) * 9525),
        heightEmu: Number(patch.heightEmu ?? Number(patch.height ?? read.properties.height ?? 0) * 9525), other });
      break;
    }
    case 'setSectionProperties': case 'setNoteProperties': {
      const document = engine.propertyDocument();
      const patch = command.kind === 'setSectionProperties' ? command.payload.patch
        : { footnotePr: command.payload.footnotePr, endnotePr: command.payload.endnotePr };
      const changed = patchSection(document, command.payload.section, patch)!;
      if (JSON.stringify(changed) !== JSON.stringify(document)) engine.setPropertyDocument(changed);
      break;
    }
    case 'setWatermark': {
      const document = engine.propertyDocument();
      const changed = watermarkDocument(document, command.payload.watermark);
      if (JSON.stringify([...changed.package.headers ?? []]) !== JSON.stringify([...document.package.headers ?? []])) engine.setPropertyDocument(changed);
      break;
    }
  }
  engine.addUndoBoundary();
  return { status: before === engine.proposalEngine.version() ? 'noop' : 'applied' };
}
