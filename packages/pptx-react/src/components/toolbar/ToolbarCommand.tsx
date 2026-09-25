import { useContext, useId, useRef } from 'react';
import type { ComponentType, CSSProperties, ReactNode } from 'react';
import type { ParagraphAlignment } from '@betteroffice/pptx';
import type { TFunction, TranslationKey } from '@betteroffice/pptx-i18n';
import type { OverflowMenuEntry } from '../../../../../shared/react-toolbar/OverflowMenu';
import {
  pptxCommandController,
  type PptxPendingCommand,
} from '../../commands/createPptxCommandStore';
import { commandShortcut, defaultArgs, isPluginCommandId } from '../../commands/descriptors';
import { MAX_FONT_POINTS, MIN_FONT_POINTS, MAX_ZOOM, MIN_ZOOM } from '../../commands/evaluate';
import { usePptxCommand, usePptxCommands, usePptxCommandState } from '../../commands/hooks';
import type {
  PptxCommandArgs,
  PptxCommandId,
  PptxCommandState,
  PptxCommandStore,
  PptxPluginCommandId,
  PptxSelectCommandId,
  PptxZOrderMove,
} from '../../commands/types';
import { useTranslation } from '../../i18n';
import { SHAPE_PRESETS, type PptxShapePreset } from '../toolbarTypes';
import { ColorPicker } from '../ui/ColorPicker';
import { EditableCombobox } from '../ui/EditableCombobox';
import { ToolbarIcon, type ToolbarIconName } from '../ui/ToolbarIcon';
import {
  ToolbarButton,
  ToolbarButtonBase,
  ToolbarDropdownBase,
  ToolbarMenuItem,
  toolbarColors,
} from '../ui/ToolbarPrimitives';
import { runFromControl } from './activation';
import { ToolbarOverflowContext, useOverflowSource, type OverflowPrompt } from './overflowRegistry';

const COMMAND_ICONS: Partial<Record<PptxCommandId, ToolbarIconName>> = {
  bold: 'bold',
  italic: 'italic',
  underline: 'underline',
  textColor: 'textColor',
  insertSlide: 'newSlide',
  insertImage: 'insertImage',
  shapeFill: 'fillColor',
  shapeStrokeColor: 'borderColor',
  shapeStrokeWidth: 'borderWidth',
  zOrder: 'bringToFront',
  save: 'save',
  exportPng: 'image',
  undo: 'undo',
  redo: 'redo',
};

const ALIGNMENTS: Record<
  ParagraphAlignment,
  { icon: ToolbarIconName; labelKey: TranslationKey; testId: string }
> = {
  l: { icon: 'alignLeft', labelKey: 'toolbar.align.left', testId: 'left' },
  ctr: { icon: 'alignCenter', labelKey: 'toolbar.align.center', testId: 'center' },
  r: { icon: 'alignRight', labelKey: 'toolbar.align.right', testId: 'right' },
  just: { icon: 'alignJustify', labelKey: 'toolbar.align.justify', testId: 'justify' },
};

const Z_ORDER: Record<PptxZOrderMove, { icon: ToolbarIconName; labelKey: TranslationKey }> = {
  forward: { icon: 'bringForward', labelKey: 'toolbar.bringForward' },
  backward: { icon: 'sendBackward', labelKey: 'toolbar.sendBackward' },
  front: { icon: 'bringToFront', labelKey: 'toolbar.bringToFront' },
  back: { icon: 'sendToBack', labelKey: 'toolbar.sendToBack' },
};

const Z_ORDER_MENU: readonly PptxZOrderMove[] = ['forward', 'backward', 'front', 'back'];

const TEST_IDS: Partial<Record<PptxCommandId, string>> = {
  bold: 'pptx-bold',
  italic: 'pptx-italic',
  underline: 'pptx-underline',
  insertSlide: 'pptx-new-slide',
  insertImage: 'pptx-insert-image',
  save: 'pptx-save',
  exportPng: 'pptx-export-png',
  undo: 'pptx-undo',
  redo: 'pptx-redo',
};

function reasonOf(state: PptxCommandState): string | undefined {
  return state.enabled ? undefined : state.disabledReason.message;
}

function record(args: unknown): Record<string, unknown> | null {
  return args && typeof args === 'object' ? (args as Record<string, unknown>) : null;
}

function iconFor(id: PptxCommandId, args: unknown): ToolbarIconName | undefined {
  const value = record(args)?.value;
  if (id === 'alignment' && typeof value === 'string')
    return ALIGNMENTS[value as ParagraphAlignment]?.icon;
  if (id === 'zOrder' && typeof value === 'string') return Z_ORDER[value as PptxZOrderMove]?.icon;
  if (id === 'tool' && typeof value === 'string') {
    return value === 'select' ? 'select' : value === 'textBox' ? 'textBox' : 'shape';
  }
  if (id === 'fontSizeStep') return record(args)?.direction === 'increase' ? 'add' : 'remove';
  return COMMAND_ICONS[id];
}

function labelFor(id: PptxCommandId, args: unknown, t: TFunction, fallback: string): string {
  const value = record(args)?.value;
  if (id === 'alignment' && typeof value === 'string') {
    const alignment = ALIGNMENTS[value as ParagraphAlignment];
    if (alignment) return t(alignment.labelKey);
  }
  if (id === 'zOrder' && typeof value === 'string') {
    const move = Z_ORDER[value as PptxZOrderMove];
    if (move) return t(move.labelKey);
  }
  if (id === 'tool' && typeof value === 'string') {
    if (value === 'select') return t('commands.selectTool');
    if (value === 'textBox') return t('toolbar.textBoxTool');
    const preset = SHAPE_PRESETS.find((candidate) => `shape:${candidate.geometry}` === value);
    if (preset) return t(preset.labelKey);
  }
  if (id === 'fontSizeStep') {
    return t(
      record(args)?.direction === 'increase'
        ? 'toolbar.increaseFontSize'
        : 'toolbar.decreaseFontSize'
    );
  }
  return fallback;
}

function testIdFor(id: PptxCommandId, args: unknown): string | undefined {
  const value = record(args)?.value;
  if (id === 'alignment' && typeof value === 'string') {
    const alignment = ALIGNMENTS[value as ParagraphAlignment];
    return alignment ? `pptx-align-${alignment.testId}` : undefined;
  }
  if (id === 'tool' && value === 'select') return 'pptx-tool-select';
  if (id === 'tool' && value === 'textBox') return 'pptx-tool-text-box';
  if (id === 'fontSizeStep') {
    return record(args)?.direction === 'increase'
      ? 'pptx-font-size-increase'
      : 'pptx-font-size-decrease';
  }
  if (id === 'insertSlide' && record(args)?.layoutPartPath !== undefined) return undefined;
  return TEST_IDS[id];
}

function hintFor(id: PptxCommandId, args: unknown): string | null {
  if (id === 'tool' && record(args)?.value === 'select') return 'Esc';
  return commandShortcut(id, args as never);
}

/** Arguments a control binds; required when the command cannot run without them. */
export type ToolbarCommandArgs<K extends PptxCommandId | PptxPluginCommandId> =
  K extends PptxCommandId
    ? null extends PptxCommandArgs[K]
      ? { args?: PptxCommandArgs[K] }
      : {} extends PptxCommandArgs[K]
      ? { args?: PptxCommandArgs[K] }
      : { args: PptxCommandArgs[K] }
    : { args?: null };

/** A built-in or contributed command; a contributed one renders nothing while inactive. */
export type ToolbarCommandButtonProps<K extends PptxCommandId | PptxPluginCommandId> = {
  id: K;
  /** Button content; defaults to the command's icon, or its label when it has none. */
  children?: ReactNode;
  /** Accessible name; defaults to the localized or contributed command label. */
  label?: string;
  className?: string;
  style?: CSSProperties;
} & ToolbarCommandArgs<K>;

export interface ToolbarCommandSelectProps<K extends PptxSelectCommandId> {
  id: K;
  className?: string;
}

/** Commands whose full built-in control chooses their arguments. */
export type PptxControlCommandId =
  | PptxSelectCommandId
  | 'textColor'
  | 'shapeFill'
  | 'shapeStrokeColor'
  | 'zOrder'
  | 'fontSizeStep';

export type ToolbarCommandProps<K extends PptxCommandId | PptxPluginCommandId> = {
  id: K;
  className?: string;
} & (K extends PptxControlCommandId
  ? {
      /** Binds arguments; omit to render the command's full built-in control. */
      args?: PptxCommandArgs[K];
    }
  : ToolbarCommandArgs<K>);

function prepare<K extends PptxCommandId>(store: PptxCommandStore, id: K): PptxPendingCommand<K> {
  return (
    pptxCommandController(store)?.prepare(id) ?? {
      execute: (args) => store.execute(id, args),
    }
  );
}

/** What overflow presentations need beyond the store. */
export interface OverflowTools {
  prompt?: (request: OverflowPrompt) => void;
}

function commandItem<K extends PptxCommandId>(
  store: PptxCommandStore,
  t: TFunction,
  id: K,
  args: PptxCommandArgs[K] | undefined,
  key: string
): OverflowMenuEntry {
  const state = store.getState(id, args) as PptxCommandState;
  const icon = iconFor(id, args);
  return {
    kind: 'item',
    id: key,
    label: labelFor(id, args, t, t(store.getDescriptor(id).labelKey)),
    icon: icon ? <ToolbarIcon name={icon} size={16} /> : undefined,
    shortcut: hintFor(id, args) ?? undefined,
    checked: state.active,
    disabled: !state.enabled,
    description: reasonOf(state),
    onSelect: () => void store.execute(id, (args ?? defaultArgs(id)) as PptxCommandArgs[K]),
  };
}

function optionEntries<K extends PptxSelectCommandId>(
  store: PptxCommandStore,
  id: K,
  state: PptxCommandState<K>
): OverflowMenuEntry[] {
  return (state.options ?? []).map((option, index) => ({
    kind: 'item',
    id: `${id}-${index}`,
    label: option.label,
    radio: true,
    checked: option.state.active === true,
    disabled: !option.state.enabled,
    description: option.state.enabled ? undefined : option.state.disabledReason.message,
    onSelect: () => void store.execute(id, option.args),
  }));
}

function selectSubmenu<K extends PptxSelectCommandId>(
  store: PptxCommandStore,
  t: TFunction,
  id: K,
  tools: OverflowTools
): OverflowMenuEntry {
  const state = store.getState(id) as PptxCommandState<K>;
  const reason = reasonOf(state as PptxCommandState);
  const entries = optionEntries(store, id, state);
  if (id === 'insertSlide') {
    entries.unshift(commandItem(store, t, 'insertSlide', {}, 'insertSlide-current'), {
      kind: 'separator',
      id: 'insertSlide-separator',
    });
  }
  if (id === 'fontSize' && tools.prompt) {
    const prompt = tools.prompt;
    entries.push({
      kind: 'item',
      id: 'fontSize-custom',
      label: t('commands.customFontSize'),
      disabled: !state.enabled,
      description: reason,
      onSelect: () => {
        const pending = prepare(store, 'fontSize');
        prompt({
          title: t('toolbar.fontSize'),
          label: t('commands.fontSizePoints'),
          initialValue: typeof state.value === 'number' ? String(state.value) : '',
          valid: (value) =>
            /^\d+(\.\d+)?$/.test(value) &&
            store.getState('fontSize', { points: Number(value) }).enabled,
          submit: (value) => void pending.execute({ points: Number(value) }),
        });
      },
    });
  }
  return {
    kind: 'submenu',
    id,
    label: t(store.getDescriptor(id).labelKey),
    disabled: !state.enabled,
    description: reason,
    entries,
  };
}

type ColorCommandId = 'textColor' | 'shapeFill' | 'shapeStrokeColor';

const CLEAR_LABELS: Partial<Record<ColorCommandId, TranslationKey>> = {
  shapeFill: 'toolbar.noFill',
  shapeStrokeColor: 'toolbar.noBorder',
};

function colorArgs(color: string | null): { color: string | null } {
  return { color };
}

function colorSubmenu(
  store: PptxCommandStore,
  t: TFunction,
  id: ColorCommandId,
  tools: OverflowTools
): OverflowMenuEntry {
  const state = store.getState(id) as PptxCommandState;
  const disabled = !state.enabled;
  const description = reasonOf(state);
  const current = typeof state.value === 'string' ? state.value : null;
  const label = t(store.getDescriptor(id).labelKey);
  const entries: OverflowMenuEntry[] = [];
  const clear = CLEAR_LABELS[id];
  if (clear) {
    entries.push({
      kind: 'item',
      id: `${id}-clear`,
      label: t(clear),
      radio: true,
      checked: state.value === null,
      disabled,
      description,
      onSelect: () => void store.execute(id, colorArgs(null) as never),
    });
  }
  if (tools.prompt) {
    const prompt = tools.prompt;
    entries.push({
      kind: 'item',
      id: `${id}-custom`,
      label: t('commands.customColor'),
      disabled,
      description,
      onSelect: () => {
        const pending = prepare(store, id);
        prompt({
          title: label,
          label: t('commands.hexColor'),
          placeholder: '#1A73E8',
          initialValue: current ?? '',
          valid: (value) => /^#?[0-9a-f]{6}$/i.test(value),
          submit: (value) =>
            void pending.execute(colorArgs(`#${value.replace(/^#/, '').toLowerCase()}`) as never),
        });
      },
    });
  }
  return { kind: 'submenu', id, label, disabled, description, entries };
}

function zOrderSubmenu(store: PptxCommandStore, t: TFunction): OverflowMenuEntry {
  const state = store.getState('zOrder');
  return {
    kind: 'submenu',
    id: 'zOrder',
    label: t(store.getDescriptor('zOrder').labelKey),
    disabled: !state.enabled,
    description: reasonOf(state as PptxCommandState),
    entries: Z_ORDER_MENU.map((value) =>
      commandItem(store, t, 'zOrder', { value }, `zOrder-${value}`)
    ),
  };
}

const SELECT_IDS: ReadonlySet<PptxCommandId> = new Set<PptxSelectCommandId>([
  'fontFamily',
  'fontSize',
  'alignment',
  'insertSlide',
  'tool',
  'shapeStrokeWidth',
  'shapeAdjustment',
  'zoom',
  'proposalSelect',
]);

/** The overflow-menu presentation of one command. */
export function commandOverflowEntry<K extends PptxCommandId>(
  store: PptxCommandStore,
  t: TFunction,
  id: K,
  args?: PptxCommandArgs[K],
  tools: OverflowTools = {}
): OverflowMenuEntry {
  if (args === undefined) {
    if (SELECT_IDS.has(id)) return selectSubmenu(store, t, id as PptxSelectCommandId, tools);
    if (id === 'textColor' || id === 'shapeFill' || id === 'shapeStrokeColor') {
      return colorSubmenu(store, t, id, tools);
    }
    if (id === 'zOrder') return zOrderSubmenu(store, t);
  }
  return commandItem(store, t, id, args, args === undefined ? id : `${id}:${JSON.stringify(args)}`);
}

function useCommandOverflow<K extends PptxCommandId>(
  element: React.RefObject<HTMLElement | null>,
  id: K,
  args?: PptxCommandArgs[K]
): void {
  const store = usePptxCommands();
  const { t } = useTranslation();
  const prompt = useContext(ToolbarOverflowContext)?.prompt;
  useOverflowSource(element, () => [commandOverflowEntry(store, t, id, args, { prompt })]);
}

/** A button bound to one command, showing its pressed and disabled state. */
export function ToolbarCommandButton<K extends PptxCommandId | PptxPluginCommandId>(
  props: ToolbarCommandButtonProps<K>
) {
  const { id, children, label, className, style } = props;
  const args = (props as { args?: unknown }).args;
  const builtIn = isPluginCommandId(id) ? null : (id as PptxCommandId);
  const command = usePptxCommand(
    id as PptxCommandId,
    (args ?? (builtIn ? defaultArgs(builtIn) : null)) as never
  );
  const { t } = useTranslation();
  if (!builtIn && command.descriptor === null) return null;
  const icon = builtIn ? iconFor(builtIn, args) : undefined;
  const name = label ?? (builtIn ? labelFor(builtIn, args, t, command.label) : command.label);
  return (
    <ToolbarButton
      active={command.state.active}
      disabled={!command.state.enabled}
      description={reasonOf(command.state as PptxCommandState)}
      title={name}
      shortcut={(builtIn ? hintFor(builtIn, args) : command.shortcut) ?? undefined}
      className={className}
      style={!icon && !children ? { padding: '0 8px', ...style } : style}
      testId={builtIn ? testIdFor(builtIn, args) : undefined}
      onClick={() => void command.execute()}
    >
      {children ?? (icon ? <ToolbarIcon name={icon} /> : name)}
    </ToolbarButton>
  );
}

function OverflowUnit({
  id,
  children,
  className,
  style,
}: {
  id: PptxCommandId;
  children: ReactNode;
  className?: string;
  style?: CSSProperties;
}) {
  const wrapperRef = useRef<HTMLSpanElement>(null);
  useCommandOverflow(wrapperRef, id);
  return (
    <span
      ref={wrapperRef}
      className={className}
      style={{ display: 'inline-flex', alignItems: 'center', gap: 1, flex: '0 0 auto', ...style }}
    >
      {children}
    </span>
  );
}

function FontFamilyMenu() {
  const store = usePptxCommands();
  const state = usePptxCommandState('fontFamily');
  const { t } = useTranslation();
  return (
    <ToolbarDropdownBase
      title={t('toolbar.fontFamily')}
      disabled={!state.enabled}
      description={reasonOf(state as PptxCommandState)}
      menuWidth={210}
      testId="pptx-font-family"
      style={{ width: 120, justifyContent: 'space-between' }}
      trigger={
        <>
          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>
            {state.value ?? t('toolbar.mixed')}
          </span>
          <ToolbarIcon name="chevronDown" size={13} />
        </>
      }
    >
      {(close) =>
        (state.options ?? []).map((option) => (
          <ToolbarMenuItem
            key={option.args.family}
            label={option.label}
            radio
            selected={state.value === option.args.family}
            onClick={() => runFromControl(store, 'fontFamily', option.args)}
            close={close}
          />
        ))
      }
    </ToolbarDropdownBase>
  );
}

function parseNumber(value: string): number {
  return Number.parseFloat(value.replace('%', '').trim());
}

function FontSizeBox() {
  const store = usePptxCommands();
  const state = usePptxCommandState('fontSize');
  const { t } = useTranslation();
  return (
    <EditableCombobox
      value={typeof state.value === 'number' ? String(state.value) : ''}
      options={(state.options ?? []).map((option) => ({
        value: String(option.args.points),
        label: option.label,
      }))}
      label={t('toolbar.fontSize')}
      disabled={!state.enabled}
      description={reasonOf(state as PptxCommandState)}
      onCommit={(value) => {
        const points = parseNumber(value);
        if (Number.isFinite(points) && points >= MIN_FONT_POINTS && points <= MAX_FONT_POINTS) {
          runFromControl(store, 'fontSize', { points });
        }
      }}
      width={50}
      inputStyle={{ textAlign: 'center' }}
      testId="pptx-font-size"
    />
  );
}

function ZoomBox() {
  const store = usePptxCommands();
  const state = usePptxCommandState('zoom');
  const { t } = useTranslation();
  const fit = t('toolbar.fit');
  const value = state.value ?? 'fit';
  const text = value === 'fit' ? fit : `${Math.round(value * 100)}%`;
  return (
    <EditableCombobox
      value={text}
      options={(state.options ?? []).map((option) => ({
        value: option.label,
        label: option.label,
      }))}
      label={t('toolbar.zoomValue', { value: text })}
      disabled={!state.enabled}
      description={reasonOf(state as PptxCommandState)}
      onCommit={(entered) => {
        if (entered === fit) {
          runFromControl(store, 'zoom', { scale: 'fit' });
          return;
        }
        const scale = parseNumber(entered) / 100;
        if (Number.isFinite(scale) && scale >= MIN_ZOOM && scale <= MAX_ZOOM) {
          runFromControl(store, 'zoom', { scale });
        }
      }}
      width={76}
      testId="pptx-zoom"
    />
  );
}

function AlignmentButtons() {
  const store = usePptxCommands();
  const state = usePptxCommandState('alignment');
  const { t } = useTranslation();
  return (
    <>
      {(state.options ?? []).map((option) => (
        <AlignmentButton key={option.args.value} store={store} value={option.args.value} t={t} />
      ))}
    </>
  );
}

function AlignmentButton({
  store,
  value,
  t,
}: {
  store: PptxCommandStore;
  value: ParagraphAlignment;
  t: TFunction;
}) {
  const state = usePptxCommandState('alignment', { value });
  const alignment = ALIGNMENTS[value];
  return (
    <ToolbarButtonBase
      title={t(alignment.labelKey)}
      active={state.active}
      disabled={!state.enabled}
      description={reasonOf(state as PptxCommandState)}
      onClick={() => void store.execute('alignment', { value })}
      testId={`pptx-align-${alignment.testId}`}
    >
      <ToolbarIcon name={alignment.icon} />
    </ToolbarButtonBase>
  );
}

function NewSlideControl() {
  const store = usePptxCommands();
  const state = usePptxCommandState('insertSlide');
  const main = usePptxCommand('insertSlide', {});
  const { t } = useTranslation();
  const layouts = state.options ?? [];
  return (
    <>
      <ToolbarButtonBase
        title={t('toolbar.newSlide')}
        disabled={!main.state.enabled}
        description={reasonOf(main.state as PptxCommandState)}
        onClick={() => void main.execute()}
        style={{ borderRadius: '4px 0 0 4px' }}
        testId="pptx-new-slide"
      >
        <ToolbarIcon name="newSlide" />
      </ToolbarButtonBase>
      <ToolbarDropdownBase
        title={t('toolbar.newSlideWithLayout')}
        disabled={!state.enabled || layouts.length === 0}
        description={reasonOf(state as PptxCommandState)}
        menuWidth={230}
        testId="pptx-new-slide-layout"
        style={{ minWidth: 20, width: 20, padding: 0, borderRadius: '0 4px 4px 0' }}
        trigger={<ToolbarIcon name="chevronDown" size={13} />}
      >
        {(close) =>
          layouts.map((layout, index) => (
            <ToolbarMenuItem
              key={layout.args.layoutPartPath ?? `default-${index}`}
              label={layout.label}
              radio
              selected={(layout.args.layoutPartPath ?? null) === (state.value ?? null)}
              onClick={() => void store.execute('insertSlide', layout.args)}
              close={close}
            />
          ))
        }
      </ToolbarDropdownBase>
    </>
  );
}

/** The shape placement tools, as a menu of presets. */
export function ShapeToolMenu() {
  const store = usePptxCommands();
  const state = usePptxCommandState('tool');
  const shape = usePptxCommandState('tool', { value: 'shape:rect' });
  const { t } = useTranslation();
  const wrapperRef = useRef<HTMLSpanElement>(null);
  const id = useId();
  useOverflowSource(wrapperRef, () => [
    {
      kind: 'submenu',
      id,
      label: t('toolbar.shapeTool'),
      disabled: !shape.enabled,
      description: reasonOf(shape as PptxCommandState),
      entries: SHAPE_PRESETS.map((preset) =>
        commandItem(
          store,
          t,
          'tool',
          { value: `shape:${preset.geometry}` },
          `${id}-${preset.geometry}`
        )
      ),
    },
  ]);
  const tool = state.value ?? 'select';
  return (
    <span ref={wrapperRef} style={{ display: 'inline-flex', flex: '0 0 auto' }}>
      <ToolbarDropdownBase
        title={t('toolbar.shapeTool')}
        active={tool.startsWith('shape:')}
        disabled={!shape.enabled}
        description={reasonOf(shape as PptxCommandState)}
        menuWidth={264}
        testId="pptx-tool-shape"
        style={{ minWidth: 46, padding: '0 4px' }}
        trigger={
          <>
            <ToolbarIcon name="shape" />
            <ToolbarIcon name="chevronDown" size={11} />
          </>
        }
      >
        {(close) => (
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(5, 44px)',
              gap: 4,
              padding: 2,
            }}
          >
            {SHAPE_PRESETS.map((preset) => {
              const value = `shape:${preset.geometry}` as const;
              return (
                <button
                  key={preset.geometry}
                  type="button"
                  role="menuitemradio"
                  tabIndex={-1}
                  aria-checked={tool === value}
                  data-testid={`pptx-shape-${preset.geometry}`}
                  aria-label={t(preset.labelKey)}
                  title={t(preset.labelKey)}
                  onClick={() => {
                    close();
                    runFromControl(store, 'tool', { value });
                  }}
                  style={{
                    appearance: 'none',
                    display: 'grid',
                    placeItems: 'center',
                    width: 44,
                    height: 38,
                    padding: 5,
                    border: `1px solid ${toolbarColors.border}`,
                    borderRadius: 4,
                    background: tool === value ? toolbarColors.active : toolbarColors.surface,
                    color: toolbarColors.text,
                    cursor: 'pointer',
                  }}
                >
                  <ShapePresetIcon geometry={preset.geometry} />
                </button>
              );
            })}
          </div>
        )}
      </ToolbarDropdownBase>
    </span>
  );
}

function StrokeWidthMenu() {
  const store = usePptxCommands();
  const state = usePptxCommandState('shapeStrokeWidth');
  const { t } = useTranslation();
  return (
    <ToolbarDropdownBase
      title={t('toolbar.borderWidth')}
      disabled={!state.enabled}
      description={reasonOf(state as PptxCommandState)}
      menuWidth={170}
      testId="pptx-shape-border-width"
      trigger={<ToolbarIcon name="borderWidth" />}
    >
      {(close) =>
        (state.options ?? []).map((option) => (
          <ToolbarMenuItem
            key={String(option.args.points)}
            label={option.label}
            radio
            selected={state.value === option.args.points}
            icon={
              option.args.points === null ? undefined : (
                <span
                  aria-hidden="true"
                  style={{
                    width: 18,
                    borderTop: `${Math.min(option.args.points, 5)}px solid currentColor`,
                  }}
                />
              )
            }
            onClick={() => void store.execute('shapeStrokeWidth', option.args)}
            close={close}
          />
        ))
      }
    </ToolbarDropdownBase>
  );
}

function AdjustmentBox() {
  const store = usePptxCommands();
  const state = usePptxCommandState('shapeAdjustment');
  const { t } = useTranslation();
  const adjustment = state.value;
  if (!adjustment) return null;
  const options = state.options ?? [];
  const limit = options.length > 0 ? options[options.length - 1].args.value : 1;
  const corner = limit < 1;
  return (
    <EditableCombobox
      value={`${Math.round(adjustment.value * 100)}%`}
      options={options.map((option) => ({
        value: String(Math.round(option.args.value * 100)),
        label: option.label,
      }))}
      label={t(corner ? 'toolbar.cornerRadius' : 'toolbar.shapeAdjustment')}
      disabled={!state.enabled}
      description={reasonOf(state as PptxCommandState)}
      onCommit={(value) => {
        const percent = parseNumber(value);
        if (!Number.isFinite(percent)) return;
        runFromControl(store, 'shapeAdjustment', {
          name: adjustment.name,
          value: Math.max(0, Math.min(limit * 100, percent)) / 100,
        });
      }}
      width={68}
      inputStyle={{ textAlign: 'center' }}
      testId={corner ? 'pptx-shape-corner-radius' : 'pptx-shape-adjustment'}
    />
  );
}

function ProposalPicker({ className }: { className?: string }) {
  const store = usePptxCommands();
  const state = usePptxCommandState('proposalSelect');
  const { t } = useTranslation();
  const disabled = !state.enabled;
  return (
    <select
      className={className}
      aria-label={t('proposals.canvasTitle')}
      aria-disabled={disabled || undefined}
      title={reasonOf(state as PptxCommandState)}
      value={state.value ?? ''}
      onChange={(event) => {
        if (!disabled) void store.execute('proposalSelect', { proposalId: event.target.value });
      }}
      style={{ maxWidth: 220, font: '400 13px ui-sans-serif, system-ui, sans-serif' }}
    >
      {(state.options ?? []).map((option) => (
        <option key={option.args.proposalId} value={option.args.proposalId}>
          {option.label}
        </option>
      ))}
    </select>
  );
}

/** The built-in picker of a selector command. */
export function ToolbarCommandSelect<K extends PptxSelectCommandId>({
  id,
  className,
}: ToolbarCommandSelectProps<K>) {
  let control: ReactNode = null;
  switch (id) {
    case 'fontFamily':
      control = <FontFamilyMenu />;
      break;
    case 'fontSize':
      control = <FontSizeBox />;
      break;
    case 'alignment':
      control = <AlignmentButtons />;
      break;
    case 'insertSlide':
      control = <NewSlideControl />;
      break;
    case 'tool':
      control = (
        <>
          <ToolButton value="select" />
          <ToolButton value="textBox" />
          <ShapeToolMenu />
        </>
      );
      break;
    case 'shapeStrokeWidth':
      control = <StrokeWidthMenu />;
      break;
    case 'shapeAdjustment':
      control = <AdjustmentBox />;
      break;
    case 'zoom':
      control = <ZoomBox />;
      break;
    case 'proposalSelect':
      control = <ProposalPicker />;
      break;
  }
  if (id === 'tool') {
    return (
      <span
        className={className}
        style={{ display: 'inline-flex', alignItems: 'center', gap: 1, flex: '0 0 auto' }}
      >
        {control}
      </span>
    );
  }
  return (
    <OverflowUnit id={id} className={className}>
      {control}
    </OverflowUnit>
  );
}

function ToolButton({ value }: { value: 'select' | 'textBox' }) {
  return <ToolbarCommandButton id="tool" args={{ value }} />;
}

const COLOR_TEST_IDS: Record<ColorCommandId, string> = {
  textColor: 'pptx-text-color',
  shapeFill: 'pptx-shape-fill',
  shapeStrokeColor: 'pptx-shape-border-color',
};

const COLOR_FALLBACKS: Record<ColorCommandId, string> = {
  textColor: '#000000',
  shapeFill: '#d9eaf7',
  shapeStrokeColor: '#202124',
};

function CommandColorPicker({ id }: { id: ColorCommandId }) {
  const store = usePptxCommands();
  const state = usePptxCommandState(id);
  const { t } = useTranslation();
  const clear = CLEAR_LABELS[id];
  const value = typeof state.value === 'string' ? state.value : null;
  const disabled = !state.enabled;
  return (
    <OverflowUnit id={id}>
      <ColorPicker
        value={value ?? COLOR_FALLBACKS[id]}
        label={t(store.getDescriptor(id).labelKey)}
        clearLabel={clear ? t(clear) : undefined}
        icon={iconFor(id, undefined)}
        none={Boolean(clear) && !value}
        disabled={disabled}
        description={reasonOf(state as PptxCommandState)}
        onChange={(color) => runFromControl(store, id, colorArgs(color) as never)}
        onClear={clear ? () => void store.execute(id, colorArgs(null) as never) : undefined}
        testId={COLOR_TEST_IDS[id]}
      />
    </OverflowUnit>
  );
}

function ArrangeMenu() {
  const store = usePptxCommands();
  const state = usePptxCommandState('zOrder');
  const { t } = useTranslation();
  return (
    <OverflowUnit id="zOrder">
      <ToolbarDropdownBase
        title={t('toolbar.arrange')}
        disabled={!state.enabled}
        description={reasonOf(state as PptxCommandState)}
        menuWidth={190}
        testId="pptx-shape-arrange"
        trigger={<ToolbarIcon name="bringToFront" />}
      >
        {(close) =>
          Z_ORDER_MENU.map((value) => (
            <ArrangeItem key={value} store={store} value={value} close={close} t={t} />
          ))
        }
      </ToolbarDropdownBase>
    </OverflowUnit>
  );
}

function ArrangeItem({
  store,
  value,
  close,
  t,
}: {
  store: PptxCommandStore;
  value: PptxZOrderMove;
  close: () => void;
  t: TFunction;
}) {
  const state = usePptxCommandState('zOrder', { value });
  const move = Z_ORDER[value];
  return (
    <ToolbarMenuItem
      label={t(move.labelKey)}
      icon={<ToolbarIcon name={move.icon} size={16} />}
      disabled={!state.enabled}
      description={reasonOf(state as PptxCommandState)}
      onClick={() => void store.execute('zOrder', { value })}
      close={close}
    />
  );
}

function FontSizeSteps() {
  return (
    <>
      <ToolbarCommandButton id="fontSizeStep" args={{ direction: 'decrease' }} />
      <ToolbarCommandButton id="fontSizeStep" args={{ direction: 'increase' }} />
    </>
  );
}

/** The built-in control of any command, as the default toolbar presents it. */
export function ToolbarCommand<K extends PptxCommandId | PptxPluginCommandId>(
  props: ToolbarCommandProps<K>
) {
  const { id, className } = props;
  if (isPluginCommandId(id)) {
    const pluginId: PptxPluginCommandId = id;
    return <ToolbarCommandButton id={pluginId} className={className} />;
  }
  const args = (props as { args?: unknown }).args;
  if (args === undefined) {
    if (SELECT_IDS.has(id)) {
      return <ToolbarCommandSelect id={id as PptxSelectCommandId} className={className} />;
    }
    switch (id) {
      case 'textColor':
      case 'shapeFill':
      case 'shapeStrokeColor':
        return <CommandColorPicker id={id} />;
      case 'zOrder':
        return <ArrangeMenu />;
      case 'fontSizeStep':
        return <FontSizeSteps />;
      default:
        break;
    }
  }
  const Button = ToolbarCommandButton as ComponentType<{
    id: PptxCommandId;
    args?: unknown;
    className?: string;
  }>;
  return <Button id={id} args={args} className={className} />;
}

function ShapePresetIcon({ geometry }: { geometry: PptxShapePreset }) {
  const path = {
    rect: 'M3 5h18v14H3Z',
    roundRect: 'M7 5h10a4 4 0 0 1 4 4v6a4 4 0 0 1-4 4H7a4 4 0 0 1-4-4V9a4 4 0 0 1 4-4Z',
    ellipse: 'M3 12a9 7 0 1 0 18 0 9 7 0 1 0-18 0',
    triangle: 'm12 4 9 16H3Z',
    rtTriangle: 'M4 4v16h16Z',
    diamond: 'm12 3 9 9-9 9-9-9Z',
    parallelogram: 'M7 5h14l-4 14H3Z',
    trapezoid: 'M7 5h10l4 14H3Z',
    pentagon: 'm12 3 9 7-4 11H7L3 10Z',
    hexagon: 'm7 4 10 0 5 8-5 8H7l-5-8Z',
    octagon: 'm7 3 10 0 4 4v10l-4 4H7l-4-4V7Z',
    star5: 'm12 2.5 2.8 6 6.5.6-5 4.3 1.6 6.4-5.7-3.4-5.7 3.4 1.6-6.4-5-4.3 6.5-.6Z',
    rightArrow: 'M3 8h11V4l7 8-7 8v-4H3Z',
    leftArrow: 'm21 8H10V4l-7 8 7 8v-4h11Z',
    upArrow: 'M8 21V10H4l8-7 8 7h-4v11Z',
    downArrow: 'M8 3v11H4l8 7 8-7h-4V3Z',
    chevron: 'M4 4h9l7 8-7 8H4l7-8Z',
  }[geometry];
  return (
    <svg
      width="30"
      height="24"
      viewBox="0 0 24 24"
      fill="rgba(60, 64, 67, 0.08)"
      stroke="currentColor"
      strokeWidth="1.4"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d={path} />
    </svg>
  );
}
