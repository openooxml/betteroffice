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

import type {
  DisplayPage,
  DisplayPrimitive,
  GlyphRunPrimitive,
  ImagePrimitive,
  InlineSdtWidgetAttrs,
  SdtAttrs,
  ShapePrimitive,
  TextRunPrimitive,
} from './displayList';
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
  primitives: DisplayPrimitive[];
}

interface ControlExtent {
  rect: GeoRect;
  occluded: boolean;
}

interface BoundaryExtent extends SdtExtent {
  kind: ReturnType<typeof blockWidgetKind>;
  widget?: ControlExtent;
  repeats: ControlExtent[];
}

interface InteractiveOverlayLayer {
  lowerLayer: boolean;
  boundaries: BoundaryExtent[];
  widgets: Array<WidgetExtent & { occluded: boolean }>;
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

  for (const [layer, controls] of collectInteractiveOverlayLayers(page).entries()) {
    for (const extent of controls.boundaries) {
      const boundary = renderBoundary(extent, doc, options.labels, controls.lowerLayer);
      boundary.style.zIndex = String(layer);
      root.appendChild(boundary);
    }

    for (const extent of controls.widgets) {
      const widget = renderInlineWidget(extent, doc, options.labels);
      widget.style.zIndex = String(layer);
      if (extent.occluded) hideOccludedControl(widget);
      root.appendChild(widget);
    }
  }
  return root;
}

export function collectInteractiveOverlayLayers(page: DisplayPage): InteractiveOverlayLayer[] {
  const layers = pagePrimitiveLayers(page);
  const body = layers[3]!;
  return layers.map((primitives, layer) => {
    const lowerLayer = layer < 3;
    const controlExtent = (rect: GeoRect): ControlExtent => ({
      rect,
      occluded: lowerLayer && bodyPaintsRectCenter(body, rect),
    });
    const boundaries = [...collectSdtExtents(primitives).values()]
      .sort(compareSdtExtents)
      .map((extent): BoundaryExtent => {
        const { kind, repeat } = boundaryControls(extent.attrs);
        const { rect } = extent;
        const right = rect.x + Math.max(1, rect.w) - BOUNDARY_CONTROL_INSET;
        const bottom = rect.y + Math.max(1, rect.h) - BOUNDARY_CONTROL_INSET;
        return {
          ...extent,
          kind,
          widget: kind
            ? controlExtent({
                x: right - BOUNDARY_WIDGET_SIZE,
                y: rect.y + BOUNDARY_CONTROL_INSET,
                w: BOUNDARY_WIDGET_SIZE,
                h: BOUNDARY_WIDGET_SIZE,
              })
            : undefined,
          repeats: repeat
            ? [0, 1].map((index) => controlExtent({
                x: right - REPEAT_BUTTON_SIZE -
                  (1 - index) * (REPEAT_BUTTON_SIZE + REPEAT_BUTTON_GAP),
                y: bottom - REPEAT_BUTTON_SIZE,
                w: REPEAT_BUTTON_SIZE,
                h: REPEAT_BUTTON_SIZE,
              }))
            : [],
        };
      });
    const widgets = [...collectWidgetExtents(primitives).values()].map((extent) => ({
      ...extent,
      occluded: lowerLayer && bodyPaintsRectCenter(body, extent.rect),
    }));
    return { lowerLayer, boundaries, widgets };
  });
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
      case 'glyphRun':
        if (!primitive.text.trim() || textPaintsNothing(primitive)) return false;
        // A turned or compressed run paints outside or short of its box.
        if ((primitive.rotationDeg ?? 0) % 360 !== 0) return false;
        if (primitive.horizontalScale !== undefined && !(primitive.horizontalScale >= 100)) {
          return false;
        }
        if (primitive.kind === 'glyphRun' && !primitive.glyphs.length) return false;
        break;
      case 'rect':
        if (!primitive.fill || primitive.fill === 'transparent' || primitive.fill === 'none') {
          return false;
        }
        break;
      case 'shape': {
        const paint = primitive.fillPaint;
        const fill = paint?.color ?? primitive.fill;
        if (paint?.kind === 'none') return false;
        if (paint?.kind === 'gradient' || paint?.kind === 'pattern') break;
        if (paint?.kind === 'picture' && (paint.pictureSrc || paint.pictureRelId)) {
          if ((paint.pictureOpacity ?? 1) <= 0) return false;
          // An inset or a crop past the source paints only part of the shape, tiled fills
          // included: past the tile cap the canvas stretches them.
          const inset = paint.pictureStretchRect;
          const sides = [inset?.left, inset?.top, inset?.right, inset?.bottom];
          if (sides.some((side) => (side ?? 0) > 0) || !cropFillsFrame(paint.pictureSrcRect)) {
            return false;
          }
          break;
        }
        if (!fill || fill === 'transparent' || fill === 'none') return false;
        break;
      }
      case 'image':
        break;
      case 'decoration': {
        // Only a solid rule or highlight fills its box; the others are segmented strokes.
        const { color, style, dashed, dotted } = primitive;
        if ((style ?? 'solid') !== 'solid' || dashed || dotted) return false;
        if (!color || color === 'transparent' || color === 'none') return false;
        break;
      }
      default:
        return false;
    }
    const painted =
      primitive.kind === 'shape'
        ? shapeFillRect(primitive)
        : primitive.kind === 'image'
          ? imagePaintRect(primitive)
          : displayPrimitiveRect(primitive);
    return (
      painted !== null &&
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

/**
 * The rectangle a shape's fill paints, when its path is exactly an axis-aligned
 * rectangle turned by a multiple of 180 degrees; any other path covers nothing.
 */
/**
 * The frame an image paints over whole, as the canvas draws it: none for a non-rectangular image,
 * a turn other than a half-turn, or a crop that leaves part of the frame bare.
 */
function imagePaintRect(image: ImagePrimitive): GeoRect | null {
  const finite = (value: number | undefined, fallback: number) =>
    value !== undefined && Number.isFinite(value) ? value : fallback;
  if (image.shapeType !== undefined && image.shapeType !== 'rect') return null;
  if (finite(image.rotationDeg, 0) % 180 !== 0 || !cropFillsFrame(image.crop)) return null;
  return {
    x: finite(image.contentFrame?.x, image.x),
    y: finite(image.contentFrame?.y, image.y),
    w: finite(image.contentFrame?.w, image.w),
    h: finite(image.contentFrame?.h, image.h),
  };
}

/** Whether a source crop draws over its whole frame: an outset side leaves a gutter. */
function cropFillsFrame(
  crop: { left?: number; top?: number; right?: number; bottom?: number } | undefined
): boolean {
  if (!crop) return true;
  const side = (value: number | undefined) => (Number.isFinite(value) ? (value as number) : 0);
  const [left, top, right, bottom] = [crop.left, crop.top, crop.right, crop.bottom].map(side);
  return (
    left >= 0 && top >= 0 && right >= 0 && bottom >= 0 && left + right < 1 && top + bottom < 1
  );
}

function shapeFillRect(shape: ShapePrimitive): GeoRect | null {
  const rotation = shape.transform?.rotation ?? 0;
  if (!Number.isFinite(rotation) || rotation % 180 !== 0) return null;
  const commands = shape.geometryPath;
  const end = commands.at(-1)?.type === 'close' ? commands.length - 1 : commands.length;
  const corners: Array<[number, number]> = [];
  for (let index = 0; index < end; index++) {
    const command = commands[index]!;
    if (command.type !== 'move' && command.type !== 'line') return null;
    if ((command.type === 'move') !== (index === 0)) return null;
    corners.push([command.x, command.y]);
  }
  const first = corners[0];
  const last = corners.at(-1);
  if (corners.length === 5 && first && last && first[0] === last[0] && first[1] === last[1]) {
    corners.pop();
  }
  if (corners.length !== 4) return null;
  const xs = [...new Set(corners.map(([x]) => x))];
  const ys = [...new Set(corners.map(([, y]) => y))];
  if (xs.length !== 2 || ys.length !== 2) return null;
  if (new Set(corners.map(([x, y]) => `${x},${y}`)).size !== 4) return null;
  const sides = corners.every(([x, y], index) => {
    const [nextX, nextY] = corners[(index + 1) % 4]!;
    return (x === nextX) !== (y === nextY);
  });
  if (!sides) return null;
  let left = Math.min(...xs);
  let top = Math.min(...ys);
  const width = Math.abs(xs[0]! - xs[1]!);
  const height = Math.abs(ys[0]! - ys[1]!);
  // The canvas turns and flips a shape about its box's center.
  const halfTurn = Math.abs(rotation % 360) === 180;
  if (Boolean(shape.transform?.flipH) !== halfTurn) left = 2 * shape.x + shape.w - left - width;
  if (Boolean(shape.transform?.flipV) !== halfTurn) top = 2 * shape.y + shape.h - top - height;
  return { x: left, y: top, w: width, h: height };
}

/** A run with no glyph fill: whether its outline paints differs by canvas path, so it covers nothing. */
function textPaintsNothing(run: TextRunPrimitive | GlyphRunPrimitive): boolean {
  return run.modernEffects?.textFill?.kind === 'none';
}

export function hideOccludedControl(element: HTMLElement): void {
  element.style.visibility = 'hidden';
  element.style.pointerEvents = 'none';
  element.tabIndex = -1;
  element.setAttribute('inert', '');
  element.setAttribute('aria-hidden', 'true');
  element.setAttribute('aria-disabled', 'true');
  if (element.tagName === 'BUTTON') (element as HTMLButtonElement).disabled = true;
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
    if (current) {
      current.rect = unionRect(current.rect, rect);
      current.primitives.push(primitive);
    } else widgets.set(key, { attrs, rect: { ...rect }, primitives: [primitive] });
  }
  return widgets;
}

function compareSdtExtents(a: SdtExtent, b: SdtExtent): number {
  return (a.attrs.depth ?? 0) - (b.attrs.depth ?? 0) || a.rect.y - b.rect.y || a.rect.x - b.rect.x;
}

function renderBoundary(
  extent: BoundaryExtent,
  doc: Document,
  labels: InteractiveOverlayLabels | undefined,
  lowerLayer: boolean
): HTMLElement {
  const { attrs, rect, kind } = extent;
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
    if (extent.widget?.occluded) hideOccludedControl(trigger);
    box.appendChild(trigger);
  }

  if (extent.repeats.length) {
    const controls = doc.createElement('div');
    controls.className = 'layout-sdt-repeat-controls';
    controls.style.pointerEvents = lowerLayer ? 'none' : 'auto';
    const buttons = [
      repeatButton(doc, attrs, 'add', '＋', labels?.addRepeatingItem, authoredName),
      repeatButton(doc, attrs, 'remove', '✕', labels?.removeRepeatingItem, authoredName),
    ];
    for (const [index, button] of buttons.entries()) {
      const control = extent.repeats[index]!;
      if (lowerLayer) button.style.pointerEvents = 'auto';
      if (control.occluded) hideOccludedControl(button);
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
  for (const layer of collectInteractiveOverlayLayers(page)) {
    for (const extent of layer.boundaries) {
      if (extent.widget && !extent.widget.occluded) return true;
      if (extent.repeats.some((control) => !control.occluded)) return true;
    }
    for (const extent of layer.widgets) {
      if (!extent.occluded && extent.attrs.locked !== true) return true;
    }
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
