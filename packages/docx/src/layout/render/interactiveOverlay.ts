/**
 * Framework-neutral interactive overlay for canvas-rendered content controls.
 *
 * The accessibility mirror must remain invisible and pointer-inert. This module
 * therefore derives a separate, visible DOM layer from the same display-list
 * metadata. React and Vue mount the returned element above a canvas page and
 * keep their existing delegated `.layout-sdt-widget` / repeat handlers.
 *
 * Security: every DOCX-derived string is assigned through `textContent`,
 * `dataset`, or `setAttribute`; no HTML strings are parsed.
 */

import type { DisplayPage, DisplayPrimitive, InlineSdtWidgetAttrs, SdtAttrs } from './displayList';
import {
  clipPaintsPoint,
  displayPrimitiveRect,
  glyphRunRect,
  lineRect,
  textRunRect,
  type GeoRect,
} from './displayListGeometry';

export interface InteractiveOverlayLabels {
  /** Accessible name for a repeating-section add button. */
  addRepeatingItem?: string;
  /** Accessible name for a repeating-section remove button. */
  removeRepeatingItem?: string;
  /** Accessible name used when a control has no authored alias or tag. */
  control?: string;
}

export interface BuildInteractiveOverlayOptions {
  /** Document used to create elements (default: global `document`). */
  document?: Document;
  /** Host-localized button names. */
  labels?: InteractiveOverlayLabels;
}

/** Apply content-control focus to the interactive overlay boxes. */
export function applyInteractiveSdtFocus(
  root: ParentNode,
  focusedGroupIds: ReadonlySet<string>
): void {
  for (const box of Array.from(
    root.querySelectorAll<HTMLElement>('.layout-canvas-sdt-box[data-sdt-group-id]')
  )) {
    box.classList.toggle('layout-sdt-focused', focusedGroupIds.has(box.dataset.sdtGroupId ?? ''));
  }
}

/** The only pointer-active elements the overlay renders. */
const INTERACTIVE_SELECTOR =
  '.layout-sdt-widget, .layout-inline-sdt-widget, .layout-sdt-repeat-btn';
const BOUNDARY_CONTROL_INSET = 2;
const BOUNDARY_WIDGET_SIZE = 18;
const REPEAT_BUTTON_SIZE = 16;
const REPEAT_BUTTON_GAP = 2;

interface SdtExtent {
  attrs: SdtAttrs;
  rect: GeoRect;
}

interface WidgetExtent {
  attrs: InlineSdtWidgetAttrs;
  rect: GeoRect;
}

/**
 * Build the visible/pointer-active overlay for one page.
 *
 * Coordinates remain page-local CSS pixels. The caller applies the same zoom
 * transform as the canvas and mirror. Boundary boxes themselves never receive
 * pointer events; only real buttons are interactive/focusable.
 */
export function buildInteractiveOverlayPage(
  page: DisplayPage,
  options: BuildInteractiveOverlayOptions = {}
): HTMLElement {
  const doc = options.document ?? document;
  const root = doc.createElement('div');
  root.className = 'layout-interactive-overlay';
  root.dataset.pageIndex = String(page.pageIndex);
  root.style.position = 'absolute';
  root.style.inset = '0';
  root.style.width = `${page.width}px`;
  root.style.height = `${page.height}px`;
  root.style.pointerEvents = 'none';

  // Focus-steal guard: a mousedown that bubbles past the overlay reaches the
  // adapters' canvas pointer routing, which would move the caret and shift
  // focus away from the hidden editor. Swallow it at the overlay root for the
  // interactive elements only — click still bubbles, so the adapters' existing
  // delegated `.layout-sdt-widget` / repeat handlers keep doing the activation.
  root.addEventListener('mousedown', (event) => {
    const target = event.target as HTMLElement | null;
    if (!target?.closest?.(INTERACTIVE_SELECTOR)) return;
    event.preventDefault();
    event.stopPropagation();
  });

  const layers = pagePrimitiveLayers(page);
  const body = layers[3]!;
  for (const [layer, primitives] of layers.entries()) {
    const groups = collectSdtExtents(primitives);
    for (const extent of [...groups.values()].sort(compareSdtExtents)) {
      const boundary = renderBoundary(extent, doc, options.labels, layer < 3 ? body : undefined);
      boundary.style.zIndex = String(layer);
      root.appendChild(boundary);
    }

    const widgets = collectWidgetExtents(primitives);
    for (const extent of widgets.values()) {
      const widget = renderInlineWidget(extent, doc, options.labels);
      widget.style.zIndex = String(layer);
      if (layer < 3) hideCoveredControl(widget, extent.rect, body);
      root.appendChild(widget);
    }
  }
  return root;
}

function bodyPaintsRectCenter(body: DisplayPrimitive[], rect: GeoRect): boolean {
  const x = rect.x + Math.max(1, rect.w) / 2;
  const y = rect.y + Math.max(1, rect.h) / 2;
  return body.some((primitive) => {
    if ('opacity' in primitive && primitive.opacity !== undefined && primitive.opacity <= 0) {
      return false;
    }
    if (primitive.clipGroup?.clip && (primitive.clipGroup.opacity ?? 1) <= 0) return false;
    switch (primitive.kind) {
      case 'text':
        if (!primitive.text) return false;
        break;
      case 'glyphRun':
        if (!primitive.glyphs.length) return false;
        break;
      case 'rect':
        if (!primitive.fill || primitive.fill === 'transparent' || primitive.fill === 'none') {
          return false;
        }
        break;
      case 'shape': {
        const paint = primitive.fillPaint;
        const fill = paint?.color ?? primitive.fill;
        if (!primitive.geometryPath.length) return false;
        if (paint?.kind === 'none') return false;
        if (paint?.kind === 'gradient' || paint?.kind === 'pattern') break;
        if (paint?.kind === 'picture' && (paint.pictureSrc || paint.pictureRelId)) break;
        if (!fill || fill === 'transparent' || fill === 'none') return false;
        break;
      }
      case 'image':
      case 'decoration':
        break;
      default:
        return false;
    }
    const painted = displayPrimitiveRect(primitive);
    return (
      Number.isFinite(painted.w) &&
      Number.isFinite(painted.h) &&
      painted.w > 0 &&
      painted.h > 0 &&
      x >= painted.x &&
      x <= painted.x + painted.w &&
      y >= painted.y &&
      y <= painted.y + painted.h &&
      clipPaintsPoint(primitive, x, y)
    );
  });
}

function hideCoveredControl(button: HTMLButtonElement, rect: GeoRect, body: DisplayPrimitive[]): void {
  if (!bodyPaintsRectCenter(body, rect)) return;
  button.style.visibility = 'hidden';
  button.style.pointerEvents = 'none';
  button.tabIndex = -1;
  button.disabled = true;
}

function pagePrimitiveLayers(page: DisplayPage): DisplayPrimitive[][] {
  const watermarkPrimitiveCount = Math.min(
    page.watermarkPrimitiveCount ?? 0,
    page.primitives.length
  );
  return [
    page.primitives.slice(0, watermarkPrimitiveCount),
    page.header?.primitives ?? [],
    page.footer?.primitives ?? [],
    page.primitives.slice(watermarkPrimitiveCount),
    ...(page.noteAreas ?? []).map((area) => [
      ...(area.separatorPrimitives ?? []),
      ...(area.primitives ?? []),
    ]),
  ];
}

function collectSdtExtents(primitives: DisplayPrimitive[]): Map<string, SdtExtent> {
  const groups = new Map<string, SdtExtent>();
  for (const primitive of primitives) {
    const rect = primitiveRect(primitive);
    if (!rect) continue;
    const path = primitive.sdtPath?.length
      ? primitive.sdtPath
      : primitive.sdt
        ? [primitive.sdt]
        : [];
    for (const attrs of path) {
      const current = groups.get(attrs.groupId);
      if (current) current.rect = unionRect(current.rect, rect);
      else groups.set(attrs.groupId, { attrs, rect: { ...rect } });
    }
  }
  return groups;
}

function collectWidgetExtents(primitives: DisplayPrimitive[]): Map<string, WidgetExtent> {
  const widgets = new Map<string, WidgetExtent>();
  for (const primitive of primitives) {
    const attrs = primitive.inlineSdtWidget;
    const rect = primitiveRect(primitive);
    if (!attrs || !rect) continue;
    const key = `${attrs.groupId}:${attrs.pos}:${attrs.controlKind ?? attrs.kind}`;
    const current = widgets.get(key);
    if (current) current.rect = unionRect(current.rect, rect);
    else widgets.set(key, { attrs, rect: { ...rect } });
  }
  return widgets;
}

function compareSdtExtents(a: SdtExtent, b: SdtExtent): number {
  return (a.attrs.depth ?? 0) - (b.attrs.depth ?? 0) || a.rect.y - b.rect.y || a.rect.x - b.rect.x;
}

function renderBoundary(
  extent: SdtExtent,
  doc: Document,
  labels: InteractiveOverlayLabels | undefined,
  body: DisplayPrimitive[] | undefined
): HTMLElement {
  const { attrs, rect } = extent;
  const right = rect.x + Math.max(1, rect.w) - BOUNDARY_CONTROL_INSET;
  const bottom = rect.y + Math.max(1, rect.h) - BOUNDARY_CONTROL_INSET;
  const box = doc.createElement('div');
  box.className = 'layout-block-sdt-box layout-canvas-sdt-box';
  stampSdtAttrs(box, attrs);
  placeAt(box, rect);
  box.style.pointerEvents = 'none';

  const authoredName = attrs.alias || attrs.tag;
  if (authoredName) {
    const chip = doc.createElement('span');
    chip.className = 'layout-block-sdt-label';
    chip.textContent = authoredName;
    box.appendChild(chip);
  }

  const { kind, repeat } = boundaryControls(attrs);
  if (kind) {
    const trigger = doc.createElement('button');
    trigger.type = 'button';
    trigger.className = 'layout-sdt-widget';
    trigger.dataset.sdtWidget = kind;
    trigger.dataset.sdtGroupId = attrs.groupId;
    if (attrs.tag) trigger.dataset.sdtTag = attrs.tag;
    if (attrs.alias) trigger.dataset.sdtAlias = attrs.alias;
    if (authoredName || labels?.control) {
      trigger.setAttribute('aria-label', authoredName || labels?.control || '');
    }
    if (kind === 'dropdown') trigger.setAttribute('aria-haspopup', 'listbox');
    if (kind === 'date') trigger.setAttribute('aria-haspopup', 'dialog');
    if (kind === 'checkbox') {
      trigger.setAttribute('role', 'checkbox');
      trigger.setAttribute('aria-checked', String(attrs.checked ?? false));
    }
    trigger.textContent =
      kind === 'dropdown' ? '▾' : kind === 'date' ? '▣' : attrs.checked ? '☒' : '☐';
    trigger.style.pointerEvents = 'auto';
    if (body) {
      trigger.style.width = `${BOUNDARY_WIDGET_SIZE}px`;
      trigger.style.height = `${BOUNDARY_WIDGET_SIZE}px`;
      trigger.style.boxSizing = 'border-box';
      hideCoveredControl(
        trigger,
        {
          x: right - BOUNDARY_WIDGET_SIZE,
          y: rect.y + BOUNDARY_CONTROL_INSET,
          w: BOUNDARY_WIDGET_SIZE,
          h: BOUNDARY_WIDGET_SIZE,
        },
        body
      );
    }
    box.appendChild(trigger);
  }

  if (repeat) {
    const controls = doc.createElement('div');
    controls.className = 'layout-sdt-repeat-controls';
    controls.style.pointerEvents = body ? 'none' : 'auto';
    const buttons = [
      repeatButton(doc, attrs, 'add', '＋', labels?.addRepeatingItem, authoredName),
      repeatButton(doc, attrs, 'remove', '✕', labels?.removeRepeatingItem, authoredName),
    ];
    for (const [index, button] of buttons.entries()) {
      if (body) {
        button.style.width = `${REPEAT_BUTTON_SIZE}px`;
        button.style.height = `${REPEAT_BUTTON_SIZE}px`;
        button.style.boxSizing = 'border-box';
        button.style.pointerEvents = 'auto';
        hideCoveredControl(
          button,
          {
            x:
              right - REPEAT_BUTTON_SIZE -
              (buttons.length - 1 - index) * (REPEAT_BUTTON_SIZE + REPEAT_BUTTON_GAP),
            y: bottom - REPEAT_BUTTON_SIZE,
            w: REPEAT_BUTTON_SIZE,
            h: REPEAT_BUTTON_SIZE,
          },
          body
        );
      }
      controls.appendChild(button);
    }
    box.appendChild(controls);
  }
  return box;
}

function repeatButton(
  doc: Document,
  attrs: SdtAttrs,
  operation: 'add' | 'remove',
  glyph: string,
  label: string | undefined,
  authoredName: string | undefined
): HTMLButtonElement {
  const button = doc.createElement('button');
  button.type = 'button';
  button.className = 'layout-sdt-repeat-btn';
  button.dataset.sdtRepeat = operation;
  button.dataset.sdtGroupId = attrs.groupId;
  if (attrs.tag) button.dataset.sdtTag = attrs.tag;
  if (label || authoredName) button.setAttribute('aria-label', label || authoredName || '');
  button.textContent = glyph;
  return button;
}

function renderInlineWidget(
  extent: WidgetExtent,
  doc: Document,
  labels: InteractiveOverlayLabels | undefined
): HTMLButtonElement {
  const { attrs, rect } = extent;
  const kind = attrs.controlKind ?? attrs.kind;
  const button = doc.createElement('button');
  button.type = 'button';
  button.className = 'layout-sdt-widget layout-inline-sdt-widget layout-canvas-inline-sdt-widget';
  button.dataset.sdtWidget = adapterWidgetKind(kind);
  button.dataset.sdtGroupId = attrs.groupId;
  button.dataset.sdtPos = String(attrs.pos);
  if (attrs.tag) button.dataset.sdtTag = attrs.tag;
  if (attrs.alias) button.dataset.sdtAlias = attrs.alias;
  if (attrs.controlId !== undefined) button.dataset.sdtControlId = String(attrs.controlId);
  if (attrs.value !== undefined) button.dataset.sdtValue = attrs.value;
  if (attrs.selectedIndex !== undefined) {
    button.dataset.sdtSelectedIndex = String(attrs.selectedIndex);
  }
  if (attrs.dateFormat) button.dataset.sdtDateFormat = attrs.dateFormat;
  if (attrs.dateLanguage) button.dataset.sdtDateLanguage = attrs.dateLanguage;
  if (attrs.listItems?.length) button.dataset.sdtListItems = JSON.stringify(attrs.listItems);
  if (attrs.alias || attrs.tag || labels?.control) {
    button.setAttribute('aria-label', attrs.alias || attrs.tag || labels?.control || '');
  }
  if (kind === 'checkbox') {
    button.setAttribute('role', 'checkbox');
    button.setAttribute('aria-checked', String(attrs.checked ?? false));
  } else if (kind === 'dropDownList' || kind === 'comboBox') {
    button.setAttribute('aria-haspopup', 'listbox');
    button.setAttribute('aria-expanded', 'false');
  } else if (kind === 'date') {
    button.setAttribute('aria-haspopup', 'dialog');
  }
  button.disabled = attrs.locked === true;
  button.textContent =
    kind === 'checkbox'
      ? attrs.checked
        ? '☒'
        : '☐'
      : kind === 'date'
        ? '▣'
        : kind === 'picture'
          ? '▧'
          : '▾';
  placeAt(button, rect);
  button.style.pointerEvents = 'auto';
  return button;
}

function adapterWidgetKind(
  kind: NonNullable<InlineSdtWidgetAttrs['controlKind']> | 'checkbox'
): string {
  if (kind === 'dropDownList' || kind === 'comboBox') return 'dropdown';
  return kind;
}

function blockWidgetKind(sdtType: string): 'checkbox' | 'dropdown' | 'date' | null {
  if (sdtType === 'checkbox') return 'checkbox';
  if (sdtType === 'dropDownList' || sdtType === 'comboBox') return 'dropdown';
  if (sdtType === 'date') return 'date';
  return null;
}

/** The buttons a content-control boundary shows: a widget trigger and repeat buttons. */
function boundaryControls(attrs: SdtAttrs): {
  kind: ReturnType<typeof blockWidgetKind>;
  repeat: boolean;
} {
  const mutable = !attrs.bound && !isLocked(attrs.lock);
  return {
    kind: mutable ? blockWidgetKind(attrs.sdtType) : null,
    repeat: Boolean(attrs.repeatingItem) && mutable,
  };
}

/** Whether `buildInteractiveOverlayPage(page)` holds a control Tab stops at. */
export function interactiveOverlayHasTabStops(page: DisplayPage): boolean {
  const primitives = pagePrimitiveLayers(page).flat();
  for (const extent of collectSdtExtents(primitives).values()) {
    const { kind, repeat } = boundaryControls(extent.attrs);
    if (kind || repeat) return true;
  }
  for (const extent of collectWidgetExtents(primitives).values()) {
    if (extent.attrs.locked !== true) return true;
  }
  return false;
}

function isLocked(lock: string | undefined): boolean {
  return lock === 'contentLocked' || lock === 'sdtContentLocked' || lock === 'sdtLocked';
}

function stampSdtAttrs(el: HTMLElement, attrs: SdtAttrs): void {
  el.dataset.sdtGroupId = attrs.groupId;
  el.dataset.sdtType = attrs.sdtType;
  if (attrs.depth !== undefined) el.dataset.sdtDepth = String(attrs.depth);
  if (attrs.tag) el.dataset.sdtTag = attrs.tag;
  if (attrs.alias) el.dataset.sdtAlias = attrs.alias;
  if (attrs.lock) el.dataset.sdtLock = attrs.lock;
  if (attrs.checked !== undefined) el.dataset.sdtChecked = String(attrs.checked);
  if (attrs.bound !== undefined) el.dataset.sdtBound = String(attrs.bound);
  if (attrs.repeatingItem !== undefined) {
    el.dataset.sdtRepeatingItem = String(attrs.repeatingItem);
  }
}

function primitiveRect(primitive: DisplayPrimitive): GeoRect | undefined {
  const clip = primitive.clipGroup?.clip;
  if (
    clip?.x !== undefined &&
    clip.y !== undefined &&
    clip.w !== undefined &&
    clip.h !== undefined
  ) {
    return { x: clip.x, y: clip.y, w: clip.w, h: clip.h };
  }
  switch (primitive.kind) {
    case 'text':
      return textRunRect(primitive);
    case 'glyphRun':
      return glyphRunRect(primitive);
    case 'rect':
      return { x: primitive.x, y: primitive.y, w: primitive.w, h: primitive.h };
    case 'line':
      return lineRect(primitive);
    case 'image':
    case 'shape':
      return { x: primitive.x, y: primitive.y, w: primitive.w, h: primitive.h };
    case 'decoration':
      return { x: primitive.x, y: primitive.y, w: primitive.w, h: primitive.h };
  }
}

function unionRect(a: GeoRect, b: GeoRect): GeoRect {
  const x = Math.min(a.x, b.x);
  const y = Math.min(a.y, b.y);
  const right = Math.max(a.x + a.w, b.x + b.w);
  const bottom = Math.max(a.y + a.h, b.y + b.h);
  return { x, y, w: right - x, h: bottom - y };
}

function placeAt(el: HTMLElement, rect: GeoRect): void {
  el.style.position = 'absolute';
  el.style.left = `${rect.x}px`;
  el.style.top = `${rect.y}px`;
  el.style.width = `${Math.max(1, rect.w)}px`;
  el.style.height = `${Math.max(1, rect.h)}px`;
}
