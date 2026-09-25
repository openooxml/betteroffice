import { useCallback, useEffect, useLayoutEffect, useMemo, useRef } from 'react';
import type { RefObject } from 'react';
import type {
  DeckSnapshot,
  ParagraphAlignment,
  PresentationHandle,
  Proposal,
  SlideDisplayList,
  TextBoxPrimitive,
  TextStylePatch,
} from '@betteroffice/pptx';
import type { TFunction, Translations } from '@betteroffice/pptx-i18n';
import type { PptxEditorTool, PptxZoom, ShapeFormattingAction } from '../components/toolbarTypes';
import { findShape } from '../interactions';
import { shapeFormattingFromShape } from '../shapeFormatting';
import {
  paragraphAlignmentFromSelection,
  selectionFormattingFromStory,
  storyFormattingFromStory,
  type EffectiveTextStyle,
} from '../textFormatting';
import {
  PptxCommandAdmissionError,
  type PptxCommandBinding,
  type PptxCommandController,
  type PptxCommandOrigin,
} from './createPptxCommandStore';
import {
  commandReason,
  evaluatePptxCommand,
  DEFAULT_FONT_FAMILIES,
  DEFAULT_FONT_SIZES,
  FALLBACK_FONT_POINTS,
  nextFontSize,
  type PptxCommandEnvironment,
  type PptxProposalEnvironment,
  type PptxShapeEnvironment,
  type PptxTextEnvironment,
} from './evaluate';
import { PptxInputFailure, PptxInputStale, type PptxInputCoordinator } from './inputCoordinator';
import type {
  PptxCommandArgs,
  PptxCommandFailureCode,
  PptxCommandId,
  PptxCommandResult,
  PptxCommandState,
} from './types';

/** Outcome of the built-in save workflow. */
export type PptxSaveOutcome =
  | 'saved'
  | 'requested'
  | 'replaced'
  | 'gesture'
  | 'input-failed'
  | 'failed';

/** Outcome of exporting the current slide as PNG. */
export type PptxExportOutcome = 'exported' | 'replaced' | 'input-failed' | 'failed';

/** The editor's actions behind the commands; each returns whether it changed anything. */
export interface PptxCommandActions {
  format(patch: TextStylePatch): boolean;
  align(value: ParagraphAlignment): boolean;
  formatShape(action: ShapeFormattingAction): boolean;
  insertSlide(layoutPartPath: string | null | undefined): boolean;
  history(direction: 'undo' | 'redo'): boolean;
  setTool(tool: PptxEditorTool): boolean;
  setZoom(zoom: PptxZoom): boolean;
  /** Opens the image picker; the chosen file is inserted as later input. */
  openPicker(): void;
  save(): Promise<PptxSaveOutcome>;
  /** Captures the slide after accepted input, then paints and downloads it. */
  exportPng(): Promise<PptxExportOutcome>;
  present(): void;
  setProposalsOpen(open: boolean): void;
  selectProposal(proposalId: string): void;
  showProposalDiff(enabled: boolean): void;
  acceptProposal(proposalId: string, force: boolean): 'accepted' | 'stale';
  rejectProposal(proposalId: string): boolean;
  reportError(error: unknown): void;
  focusEditor(): void;
}

/** Everything the command binding reads from the editor, refreshed every render. */
export interface PptxCommandInputs {
  handle: PresentationHandle | null;
  handleRef: RefObject<PresentationHandle | null>;
  modelRef: RefObject<{
    snapshot: DeckSnapshot;
    slideIndex: number;
    frame: SlideDisplayList | null;
  } | null>;
  selectionRef: RefObject<{
    shapeId: string;
    storyId: string;
    anchor: number;
    focus: number;
  } | null>;
  shapeSelectionRef: RefObject<{ slideId: string; shapeId: string } | null>;
  textStyleRef: RefObject<EffectiveTextStyle>;
  activeToolRef: RefObject<PptxEditorTool>;
  generationRef: RefObject<number>;
  coordinator: PptxInputCoordinator;
  status: PptxCommandEnvironment['status'];
  readOnlyRef: RefObject<boolean>;
  zoomRef: RefObject<PptxZoom>;
  proposalsRef: RefObject<readonly Proposal[]>;
  proposalsOpenRef: RefObject<boolean>;
  /** Whether the canvas shows the selected proposal's diff. */
  reviewEnabledRef: RefObject<boolean>;
  /** The proposal chosen for the canvas; the first one on the slide when absent. */
  reviewSelectedIdRef: RefObject<string | null>;
  proposalsAvailable: boolean;
  /** A pointer gesture has not finished. */
  gestureActive(): boolean;
  /** Fallback style of text without explicit formatting. */
  baseStyle: EffectiveTextStyle;
  i18n: Translations | undefined;
  t: TFunction;
  actions: PptxCommandActions;
  /** Values whose change re-evaluates command state. */
  stamp: readonly unknown[];
}

interface EditorOrigin extends PptxCommandOrigin {
  id: PptxCommandId;
  target: string | null;
}

const IMMEDIATE: ReadonlySet<PptxCommandId> = new Set<PptxCommandId>([
  'insertImage',
  'save',
  'exportPng',
  'slideshow',
  'zoom',
  'proposalsPanel',
  'proposalSelect',
  'proposalDiff',
]);

const TEXT_COMMANDS: ReadonlySet<PptxCommandId> = new Set<PptxCommandId>([
  'bold',
  'italic',
  'underline',
  'fontFamily',
  'fontSize',
  'fontSizeStep',
  'textColor',
  'alignment',
]);

const SHAPE_COMMANDS: ReadonlySet<PptxCommandId> = new Set<PptxCommandId>([
  'shapeFill',
  'shapeStrokeColor',
  'shapeStrokeWidth',
  'shapeAdjustment',
  'zOrder',
]);

const SLIDE_COMMANDS: ReadonlySet<PptxCommandId> = new Set<PptxCommandId>(['insertSlide', 'tool']);

function hexColor(value: string | null | undefined): string | null {
  if (!value) return null;
  const rgb = value.replace(/^#/, '');
  return /^[0-9a-f]{6}$/i.test(rgb) ? `#${rgb.toLowerCase()}` : null;
}

function mixed(value: boolean | undefined): boolean | 'mixed' {
  return value === undefined ? 'mixed' : value;
}

function proposalLabel(proposal: Proposal, index: number): string {
  return `${index + 1}. ${proposal.agentId}${proposal.note ? ` · ${proposal.note}` : ''}`;
}

function proposalView(proposal: Proposal, index: number): PptxProposalEnvironment {
  return {
    id: proposal.id,
    label: proposalLabel(proposal, index),
    stale: proposal.staleTargets.length > 0,
  };
}

function executed(changed: boolean): PptxCommandResult {
  return { ok: true, status: changed ? 'executed' : 'noop' };
}

const OPENED: PptxCommandResult = { ok: true, status: 'opened' };

function currentSlide(inputs: PptxCommandInputs) {
  const model = inputs.modelRef.current;
  return model?.snapshot.slides[model.slideIndex] ?? null;
}

function selectedShape(inputs: PptxCommandInputs) {
  const slide = currentSlide(inputs);
  const selection = inputs.shapeSelectionRef.current;
  if (!slide || !selection || slide.id !== selection.slideId) return null;
  return findShape(slide.shapes, selection.shapeId);
}

function textEnvironment(inputs: PptxCommandInputs): PptxTextEnvironment | 'unsupported' | null {
  const handle = inputs.handleRef.current;
  const selection = inputs.selectionRef.current;
  const shapeStoryId = selection ? null : selectedShape(inputs)?.textStories[0]?.id ?? null;
  const storyId = selection?.storyId ?? shapeStoryId;
  if (!handle || !storyId) return null;
  try {
    const story = handle.story(storyId);
    const textBox = inputs.modelRef.current?.frame?.primitives.find(
      (primitive): primitive is TextBoxPrimitive =>
        primitive.kind === 'textBox' && primitive.storyId === storyId
    );
    const start = selection ? selection.anchor : 0;
    const end = selection ? selection.focus : story.length;
    const alignment = paragraphAlignmentFromSelection(story, textBox, start, end) ?? null;
    if (selection && selection.anchor === selection.focus) {
      const style = inputs.textStyleRef.current;
      return {
        kind: 'caret',
        bold: style.bold,
        italic: style.italic,
        underline: style.underline !== 'none',
        fontFamily: style.fontFamily,
        fontSize: style.fontSizePt,
        color: hexColor(style.color),
        alignment,
      };
    }
    const formatting = selection
      ? selectionFormattingFromStory(story, selection.anchor, selection.focus, inputs.baseStyle)
      : storyFormattingFromStory(story, inputs.baseStyle);
    return {
      kind: selection ? 'range' : 'shape',
      bold: mixed(formatting.bold),
      italic: mixed(formatting.italic),
      underline: mixed(formatting.underline),
      fontFamily: formatting.fontFamily ?? null,
      fontSize: formatting.fontSize ?? null,
      color: hexColor(formatting.textColor),
      alignment,
    };
  } catch {
    return 'unsupported';
  }
}

function shapeEnvironment(inputs: PptxCommandInputs): PptxShapeEnvironment | null {
  const slide = currentSlide(inputs);
  const shape = selectedShape(inputs);
  if (!slide || !shape) return null;
  const formatting = shapeFormattingFromShape(shape);
  const index = slide.shapes.findIndex((candidate) => candidate.id === shape.id);
  return {
    formattable: shape.kind === 'shape',
    geometry: shape.geometry ?? null,
    fill: formatting.fillColor ?? null,
    stroke: formatting.strokeColor ?? null,
    strokeWidth: formatting.strokeWidthPt ?? null,
    adjustments: shape.adjustValues ?? {},
    order: index >= 0 ? { index, count: slide.shapes.length } : null,
  };
}

function targetKey(inputs: PptxCommandInputs, id: PptxCommandId): string | null {
  if (TEXT_COMMANDS.has(id)) {
    const selection = inputs.selectionRef.current;
    return JSON.stringify(
      selection
        ? [selection.shapeId, selection.storyId, selection.anchor, selection.focus]
        : inputs.shapeSelectionRef.current
    );
  }
  if (SHAPE_COMMANDS.has(id)) return JSON.stringify(inputs.shapeSelectionRef.current);
  if (SLIDE_COMMANDS.has(id)) return currentSlide(inputs)?.id ?? null;
  return null;
}

/** Binds an editor's command store to its live state; returns the live gate. */
export function usePptxCommandBinding(
  controller: PptxCommandController,
  inputs: PptxCommandInputs
): <K extends PptxCommandId>(id: K, args?: PptxCommandArgs[K]) => PptxCommandState<K> {
  const latest = useRef(inputs);
  latest.current = inputs;

  const binding = useMemo<PptxCommandBinding>(() => {
    const environment = (executing: boolean): PptxCommandEnvironment => {
      const current = latest.current;
      const handle = current.handleRef.current;
      const ready = current.status === 'ready' && handle !== null;
      const slide = currentSlide(current);
      const model = current.modelRef.current;
      let text: PptxCommandEnvironment['text'] | undefined;
      let shape: PptxCommandEnvironment['shape'] | undefined;
      const layouts = [
        ...new Set(model?.snapshot.slides.map((candidate) => candidate.layoutPartPath) ?? []),
      ];
      const readOnly = current.readOnlyRef.current;
      const proposals = current.proposalsRef.current;
      const available = readOnly
        ? []
        : proposals.filter((proposal) =>
            proposal.changes.some((change) => change.slideId === slide?.id)
          );
      const selected =
        available.find((proposal) => proposal.id === current.reviewSelectedIdRef.current) ??
        available[0] ??
        null;
      const reviewEnabled = current.reviewEnabledRef.current;
      return {
        status: ready ? 'ready' : current.status === 'ready' ? 'loading' : current.status,
        readOnly,
        reviewing: reviewEnabled && selected !== null,
        pendingInput: !executing && current.coordinator.busy(),
        canUndo: ready && handle.canUndo(),
        canRedo: ready && handle.canRedo(),
        slide: slide ? { id: slide.id, layoutPartPath: slide.layoutPartPath } : null,
        layouts: layouts.map((layoutPartPath, index) => ({
          args: { layoutPartPath },
          label: current.t('toolbar.layoutOption', { number: index + 1 }),
        })),
        get text() {
          if (text === undefined) text = ready ? textEnvironment(current) : null;
          return text;
        },
        get shape() {
          if (shape === undefined) shape = ready ? shapeEnvironment(current) : null;
          return shape;
        },
        tool: current.activeToolRef.current,
        zoom: current.zoomRef.current,
        fontFamilies: DEFAULT_FONT_FAMILIES,
        fontSizes: DEFAULT_FONT_SIZES,
        proposals: current.proposalsAvailable
          ? {
              open: current.proposalsOpenRef.current,
              pending: proposals.map(proposalView),
              canvas: {
                available: available.map(proposalView),
                selectedId: selected?.id ?? null,
                diff: reviewEnabled,
              },
            }
          : null,
        translate: current.t,
      };
    };

    const fail = (
      code: PptxCommandFailureCode,
      env: PptxCommandEnvironment
    ): PptxCommandResult => ({
      ok: false,
      failure: commandReason(code, env),
    });

    const perform = <K extends PptxCommandId>(
      id: K,
      rawArgs: PptxCommandArgs[K],
      env: PptxCommandEnvironment
    ): PptxCommandResult | Promise<PptxCommandResult> => {
      const { actions } = latest.current;
      const args = rawArgs as Record<string, never>;
      const text = env.text && env.text !== 'unsupported' ? env.text : null;
      const guarded = (change: () => boolean): PptxCommandResult => {
        try {
          return executed(change());
        } catch (error) {
          actions.reportError(error);
          return fail('command-failed', env);
        }
      };
      switch (id) {
        case 'bold':
          return guarded(() => actions.format({ bold: text?.bold !== true }));
        case 'italic':
          return guarded(() => actions.format({ italic: text?.italic !== true }));
        case 'underline':
          return guarded(() =>
            actions.format({ underline: text?.underline === true ? 'none' : 'sng' })
          );
        case 'fontFamily':
          return guarded(() => actions.format({ fontFamily: args.family }));
        case 'fontSize':
          return guarded(() => actions.format({ fontSizePt: args.points }));
        case 'fontSizeStep':
          return guarded(() =>
            actions.format({
              fontSizePt: nextFontSize(
                text?.fontSize ?? FALLBACK_FONT_POINTS,
                env.fontSizes,
                args.direction === 'increase' ? 1 : -1
              ),
            })
          );
        case 'textColor':
          return guarded(() => actions.format({ color: args.color }));
        case 'alignment':
          return guarded(() => actions.align(args.value));
        case 'insertSlide':
          return guarded(() => actions.insertSlide(args.layoutPartPath));
        case 'insertImage':
          actions.openPicker();
          return OPENED;
        case 'tool':
          return executed(actions.setTool(args.value));
        case 'shapeFill':
          return guarded(() => actions.formatShape({ type: 'fillColor', value: args.color }));
        case 'shapeStrokeColor':
          return guarded(() => actions.formatShape({ type: 'strokeColor', value: args.color }));
        case 'shapeStrokeWidth':
          return guarded(() => actions.formatShape({ type: 'strokeWidth', value: args.points }));
        case 'shapeAdjustment':
          return guarded(() =>
            actions.formatShape({ type: 'adjust', name: args.name, value: args.value })
          );
        case 'zOrder':
          return guarded(() => actions.formatShape({ type: 'zOrder', value: args.value }));
        case 'save':
          return actions.save().then((outcome): PptxCommandResult => {
            switch (outcome) {
              case 'saved':
                return executed(true);
              case 'requested':
                return { ok: true, status: 'requested' };
              case 'replaced':
                return fail('document-replaced', env);
              case 'gesture':
                return fail('gesture-active', env);
              case 'input-failed':
                return fail('input-failed', env);
              default:
                return fail('command-failed', env);
            }
          });
        case 'exportPng':
          return actions.exportPng().then((outcome): PptxCommandResult => {
            switch (outcome) {
              case 'exported':
                return executed(true);
              case 'replaced':
                return fail('document-replaced', env);
              case 'input-failed':
                return fail('input-failed', env);
              default:
                return fail('command-failed', env);
            }
          });
        case 'slideshow':
          actions.present();
          return OPENED;
        case 'undo':
        case 'redo':
          return guarded(() => actions.history(id));
        case 'zoom':
          return executed(actions.setZoom(args.scale));
        case 'proposalsPanel': {
          const open = rawArgs === null ? !env.proposals!.open : Boolean(args.open);
          if (open === env.proposals!.open) return executed(false);
          actions.setProposalsOpen(open);
          return open ? OPENED : executed(true);
        }
        case 'proposalSelect':
          actions.selectProposal(args.proposalId);
          return executed(true);
        case 'proposalDiff':
          if (Boolean(args.enabled) === env.proposals?.canvas.diff) return executed(false);
          actions.showProposalDiff(Boolean(args.enabled));
          return executed(true);
        case 'proposalAccept': {
          try {
            return actions.acceptProposal(args.proposalId, Boolean(args.force)) === 'accepted'
              ? executed(true)
              : fail('proposal-stale', env);
          } catch (error) {
            actions.reportError(error);
            return fail('command-failed', env);
          }
        }
        case 'proposalReject':
          return guarded(() => actions.rejectProposal(args.proposalId));
        default:
          return fail('unsupported-command', env);
      }
    };

    return {
      environment,
      ordered: (id) => !IMMEDIATE.has(id),
      admit(operation) {
        return latest.current.coordinator.run(operation).catch((error: unknown) => {
          if (error instanceof PptxInputFailure) {
            throw new PptxCommandAdmissionError('input-failed');
          }
          if (error instanceof PptxInputStale) {
            throw new PptxCommandAdmissionError('document-replaced');
          }
          throw error;
        });
      },
      perform,
      capture(id): EditorOrigin | null {
        const current = latest.current;
        if (!current.handleRef.current) return null;
        return {
          id,
          generation: current.generationRef.current,
          target: current.coordinator.keyboardQueued() ? null : targetKey(current, id),
        };
      },
      resume(origin) {
        const current = latest.current;
        const { id, generation, target } = origin as EditorOrigin;
        if (!current.handleRef.current) return 'editor-unavailable';
        if (generation !== current.generationRef.current) return 'document-replaced';
        if (target !== null && target !== targetKey(current, id)) return 'target-changed';
        if (id !== 'tool' && !IMMEDIATE.has(id) && current.gestureActive()) return 'gesture-active';
        return null;
      },
      chrome: () => ({ i18n: latest.current.i18n }),
      focusEditor: () => latest.current.actions.focusEditor(),
    };
  }, []);

  useLayoutEffect(() => {
    controller.attach(binding);
    return () => controller.detach(binding);
  }, [binding, controller]);

  useLayoutEffect(() => {
    controller.refresh();
  }, inputs.stamp);

  const { handle } = inputs;
  useEffect(() => {
    if (!handle) return;
    let scheduled = false;
    return handle.onUpdate(() => {
      if (scheduled) return;
      scheduled = true;
      queueMicrotask(() => {
        scheduled = false;
        controller.refresh();
      });
    });
  }, [handle, controller]);

  return useCallback(
    <K extends PptxCommandId>(id: K, args?: PptxCommandArgs[K]) =>
      evaluatePptxCommand(id, args, binding.environment(true)),
    [binding]
  );
}
