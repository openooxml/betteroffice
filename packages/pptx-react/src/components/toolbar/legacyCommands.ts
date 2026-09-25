import { useEffect, useLayoutEffect, useMemo, useReducer, useRef, useState } from 'react';
import type { TFunction } from '@betteroffice/pptx-i18n';
import {
  createPptxCommandController,
  type PptxCommandBinding,
} from '../../commands/createPptxCommandStore';
import {
  DEFAULT_FONT_FAMILIES,
  DEFAULT_FONT_SIZES,
  FALLBACK_FONT_POINTS,
  nextFontSize,
  type PptxCommandEnvironment,
} from '../../commands/evaluate';
import type {
  PptxCommandArgs,
  PptxCommandId,
  PptxCommandResult,
  PptxCommandStore,
} from '../../commands/types';
import { useTranslation } from '../../i18n';
import type { ToolbarProps } from '../Toolbar';
import type { FormattingAction, ShapeFormattingAction } from '../toolbarTypes';

type Handler = keyof Pick<
  ToolbarProps,
  | 'onFormat'
  | 'onShapeFormat'
  | 'onInsertSlide'
  | 'onInsertImage'
  | 'onSave'
  | 'onExportPng'
  | 'onUndo'
  | 'onRedo'
  | 'onZoomChange'
  | 'onToolChange'
>;

const HANDLERS: Partial<Record<PptxCommandId, Handler>> = {
  bold: 'onFormat',
  italic: 'onFormat',
  underline: 'onFormat',
  fontFamily: 'onFormat',
  fontSize: 'onFormat',
  fontSizeStep: 'onFormat',
  textColor: 'onFormat',
  alignment: 'onFormat',
  shapeFill: 'onShapeFormat',
  shapeStrokeColor: 'onShapeFormat',
  shapeStrokeWidth: 'onShapeFormat',
  shapeAdjustment: 'onShapeFormat',
  zOrder: 'onShapeFormat',
  insertSlide: 'onInsertSlide',
  insertImage: 'onInsertImage',
  tool: 'onToolChange',
  save: 'onSave',
  exportPng: 'onExportPng',
  undo: 'onUndo',
  redo: 'onRedo',
  zoom: 'onZoomChange',
};

const REQUESTED: PptxCommandResult = { ok: true, status: 'requested' };

function legacyEnvironment(props: ToolbarProps, t: TFunction): PptxCommandEnvironment {
  const formatting = props.currentFormatting ?? {};
  const shape = props.currentShapeFormatting ?? {};
  return {
    status: 'ready',
    readOnly: false,
    reviewing: false,
    pendingInput: false,
    canUndo: props.canUndo ?? false,
    canRedo: props.canRedo ?? false,
    slide: { id: '', layoutPartPath: props.currentLayoutPartPath ?? null },
    layouts: (props.slideLayouts ?? []).map((layout, index) => ({
      args: { layoutPartPath: layout.partPath },
      label: layout.label ?? t('toolbar.layoutOption', { number: index + 1 }),
    })),
    text: props.textSelectionActive
      ? {
          kind: 'range',
          bold: formatting.bold ?? false,
          italic: formatting.italic ?? false,
          underline: formatting.underline ?? false,
          fontFamily: formatting.fontFamily ?? null,
          fontSize: formatting.fontSize ?? null,
          color: formatting.textColor ?? null,
          alignment: formatting.align ?? null,
        }
      : null,
    shape:
      props.shapeSelectionActive || props.shapeArrangeActive
        ? {
            formattable: Boolean(props.shapeSelectionActive),
            geometry: shape.geometry ?? null,
            fill: shape.fillColor ?? null,
            stroke: shape.strokeColor ?? null,
            strokeWidth: shape.strokeWidthPt ?? null,
            adjustments: shape.adjustments ?? {},
            order: null,
          }
        : null,
    tool: props.activeTool ?? 'select',
    zoom: props.zoom ?? 'fit',
    fontFamilies: props.fontFamilies ?? DEFAULT_FONT_FAMILIES,
    fontSizes: props.fontSizes ?? DEFAULT_FONT_SIZES,
    proposals: null,
    hostDisabled: (id) => {
      const handler = HANDLERS[id];
      return Boolean(props.disabled) || !handler || !props[handler];
    },
    translate: t,
  };
}

function formatAction<K extends PptxCommandId>(
  id: K,
  rawArgs: PptxCommandArgs[K],
  env: PptxCommandEnvironment
): FormattingAction | null {
  const args = rawArgs as Record<string, never>;
  switch (id) {
    case 'bold':
    case 'italic':
    case 'underline':
      return id;
    case 'fontFamily':
      return { type: 'fontFamily', value: args.family };
    case 'fontSize':
      return { type: 'fontSize', value: args.points };
    case 'fontSizeStep': {
      const current = env.text && env.text !== 'unsupported' ? env.text.fontSize : null;
      return {
        type: 'fontSize',
        value: nextFontSize(
          current ?? FALLBACK_FONT_POINTS,
          env.fontSizes,
          args.direction === 'increase' ? 1 : -1
        ),
      };
    }
    case 'textColor':
      return { type: 'textColor', value: args.color };
    case 'alignment':
      return { type: 'align', value: args.value };
    default:
      return null;
  }
}

function shapeAction<K extends PptxCommandId>(
  id: K,
  rawArgs: PptxCommandArgs[K]
): ShapeFormattingAction | null {
  const args = rawArgs as Record<string, never>;
  switch (id) {
    case 'shapeFill':
      return { type: 'fillColor', value: args.color };
    case 'shapeStrokeColor':
      return { type: 'strokeColor', value: args.color };
    case 'shapeStrokeWidth':
      return { type: 'strokeWidth', value: args.points };
    case 'shapeAdjustment':
      return { type: 'adjust', name: args.name, value: args.value };
    case 'zOrder':
      return { type: 'zOrder', value: args.value };
    default:
      return null;
  }
}

function performLegacy<K extends PptxCommandId>(
  props: ToolbarProps,
  id: K,
  rawArgs: PptxCommandArgs[K],
  env: PptxCommandEnvironment
): PptxCommandResult {
  const format = formatAction(id, rawArgs, env);
  if (format) {
    props.onFormat?.(format);
    return REQUESTED;
  }
  const shape = shapeAction(id, rawArgs);
  if (shape) {
    props.onShapeFormat?.(shape);
    return REQUESTED;
  }
  const args = rawArgs as Record<string, never>;
  switch (id) {
    case 'insertSlide':
      props.onInsertSlide?.(args.layoutPartPath);
      break;
    case 'insertImage':
      props.onInsertImage?.();
      break;
    case 'tool':
      props.onToolChange?.(args.value);
      break;
    case 'save':
      props.onSave?.();
      break;
    case 'exportPng':
      props.onExportPng?.();
      break;
    case 'undo':
      props.onUndo?.();
      break;
    case 'redo':
      props.onRedo?.();
      break;
    case 'zoom':
      props.onZoomChange?.(args.scale);
      break;
  }
  return REQUESTED;
}

/** A command store over the legacy prop-based toolbar, dispatching to its callbacks. */
export function useLegacyToolbarStore(props: ToolbarProps): PptxCommandStore {
  const { t } = useTranslation();
  const latest = useRef({ props, t });
  latest.current = { props, t };
  const [controller] = useState(() => {
    const created = createPptxCommandController();
    const binding: PptxCommandBinding = {
      environment: () => legacyEnvironment(latest.current.props, latest.current.t),
      ordered: () => false,
      admit: (operation) => Promise.resolve().then(operation),
      perform: (id, args, env) => performLegacy(latest.current.props, id, args, env),
      capture: () => ({ generation: 0 }),
      resume: () => null,
      chrome: () => ({ i18n: undefined }),
      focusEditor: () => {},
    };
    created.attach(binding);
    return created;
  });
  useLayoutEffect(() => {
    controller.refresh();
  });
  return controller.store;
}

function pressed(active: boolean | 'mixed' | undefined): boolean | undefined {
  return active === 'mixed' ? undefined : active;
}

function defined<T extends object>(value: T): T {
  const result = {} as T;
  for (const key of Object.keys(value) as Array<keyof T>) {
    if (value[key] !== undefined) result[key] = value[key];
  }
  return result;
}

function project(store: PptxCommandStore): ToolbarProps {
  const state = <K extends PptxCommandId>(id: K) => store.getState(id);
  const run = <K extends PptxCommandId>(id: K, args: PptxCommandArgs[K]) =>
    void store.execute(id, args);
  const fontFamily = state('fontFamily');
  const fontSize = state('fontSize');
  const textColor = state('textColor');
  const alignment = state('alignment');
  const adjustment = state('shapeAdjustment').value;
  const insertSlide = state('insertSlide');
  return {
    currentFormatting: defined({
      bold: pressed(state('bold').active),
      italic: pressed(state('italic').active),
      underline: pressed(state('underline').active),
      fontFamily: fontFamily.value ?? undefined,
      fontSize: fontSize.value ?? undefined,
      textColor: textColor.value ?? undefined,
      align: alignment.value ?? undefined,
    }),
    textSelectionActive: state('bold').enabled,
    onFormat: (action) => {
      if (typeof action === 'string') run(action, null);
      else if (action.type === 'fontFamily') run('fontFamily', { family: action.value });
      else if (action.type === 'fontSize') run('fontSize', { points: action.value });
      else if (action.type === 'textColor') run('textColor', { color: action.value });
      else run('alignment', { value: action.value });
    },
    currentShapeFormatting: defined({
      fillColor: state('shapeFill').value,
      strokeColor: state('shapeStrokeColor').value,
      strokeWidthPt: state('shapeStrokeWidth').value,
      adjustments: adjustment ? { [adjustment.name]: adjustment.value } : undefined,
    }),
    shapeSelectionActive: state('shapeFill').enabled,
    shapeArrangeActive: state('zOrder').enabled,
    onShapeFormat: (action) => {
      if (action.type === 'fillColor') run('shapeFill', { color: action.value });
      else if (action.type === 'strokeColor') run('shapeStrokeColor', { color: action.value });
      else if (action.type === 'strokeWidth') run('shapeStrokeWidth', { points: action.value });
      else if (action.type === 'adjust')
        run('shapeAdjustment', { name: action.name, value: action.value });
      else run('zOrder', { value: action.value });
    },
    onInsertSlide: (layoutPartPath) =>
      run('insertSlide', layoutPartPath === undefined ? {} : { layoutPartPath }),
    onInsertImage: () => run('insertImage', null),
    slideLayouts: (insertSlide.options ?? []).map((option) => ({
      partPath: option.args.layoutPartPath ?? null,
      label: option.label,
    })),
    currentLayoutPartPath: insertSlide.value ?? null,
    onSave: () => run('save', null),
    onExportPng: () => run('exportPng', null),
    onUndo: () => run('undo', null),
    onRedo: () => run('redo', null),
    canUndo: state('undo').enabled,
    canRedo: state('redo').enabled,
    zoom: state('zoom').value ?? 'fit',
    onZoomChange: (zoom) => run('zoom', { scale: zoom }),
    activeTool: state('tool').value ?? 'select',
    onToolChange: (tool) => run('tool', { value: tool }),
    fontFamilies: (fontFamily.options ?? []).map((option) => option.args.family),
    fontSizes: (fontSize.options ?? []).map((option) => option.args.points),
    disabled: !state('save').enabled,
  };
}

/**
 * The published prop-bag view of a command store, for hosts that read
 * `useEditorToolbar()` inside a command-mode toolbar. Its callbacks dispatch
 * through the store; mixed selections read as `undefined`.
 */
export function useLegacyProjection(store: PptxCommandStore): ToolbarProps {
  const [version, bump] = useReducer((value: number) => value + 1, 0);
  useEffect(() => store.subscribe(bump), [store]);
  return useMemo(() => project(store), [store, version]);
}
