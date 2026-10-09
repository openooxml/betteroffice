import { useContext, useRef } from 'react';
import type { ComponentType, ReactNode, RefObject } from 'react';
import type { TFunction } from '@betteroffice/xlsx-i18n';
import type { OverflowMenuEntry } from '../../../../../shared/react-toolbar/OverflowMenu';
import {
  xlsxCommandController,
  type XlsxPendingCommand,
} from '../../commands/createXlsxCommandStore';
import { commandLabelKey, commandShortcut, isPluginCommandId } from '../../commands/descriptors';
import {
  isHexColor,
  MAX_FONT_POINTS,
  MAX_ZOOM,
  MIN_FONT_POINTS,
  MIN_ZOOM,
} from '../../commands/evaluate';
import { useXlsxCommand, useXlsxCommands, useXlsxCommandState } from '../../commands/hooks';
import type {
  XlsxCommandArgs,
  XlsxCommandId,
  XlsxCommandResult,
  XlsxCommandState,
  XlsxCommandStore,
  XlsxPluginCommandId,
  XlsxSelectCommandId,
} from '../../commands/types';
import { useTranslation } from '../../i18n';
import type { BorderPreset, BorderStyle, HorizontalAlignment, VerticalAlignment } from '../Toolbar';
import { ColorPicker } from '../ui/ColorPicker';
import { EditableCombobox } from '../ui/EditableCombobox';
import { ToolbarIcon, type ToolbarIconName } from '../ui/ToolbarIcon';
import {
  ToolbarButtonBase,
  ToolbarDropdown,
  ToolbarMenuItem,
  ToolbarMenuSeparator,
} from '../ui/ToolbarPrimitives';
import {
  BorderGlyph,
  BorderStyleGlyph,
  ColorSwatch,
  HorizontalAlignmentGlyph,
  VerticalAlignmentGlyph,
} from './glyphs';
import { ToolbarOverflowContext, useOverflowSource, type OverflowRegistry } from './overflowRegistry';

const COMMAND_ICONS: Partial<Record<XlsxCommandId, ToolbarIconName>> = {
  searchMenus: 'search',
  undo: 'undo',
  redo: 'redo',
  print: 'print',
  paintFormat: 'formatPaint',
  bold: 'bold',
  italic: 'italic',
  strikethrough: 'strikethrough',
  merge: 'merge',
  save: 'save',
  exportPng: 'image',
  proposalsPanel: 'proposals',
};

const TEST_IDS: Partial<Record<XlsxCommandId, string>> = {
  searchMenus: 'xlsx-search-menus',
  undo: 'xlsx-undo',
  redo: 'xlsx-redo',
  save: 'xlsx-save',
  exportPng: 'xlsx-export-png',
  proposalsPanel: 'xlsx-proposals-button',
};

/** Commands whose button shows a pressed state; bound choices also show whether they apply. */
const TOGGLES: ReadonlySet<XlsxCommandId> = new Set<XlsxCommandId>([
  'bold',
  'italic',
  'strikethrough',
  'paintFormat',
  'proposalsPanel',
  'borderStyle',
  'horizontalAlignment',
  'verticalAlignment',
  'textWrapping',
]);

const SELECT_COMMANDS: ReadonlySet<XlsxCommandId> = new Set<XlsxSelectCommandId>([
  'fontFamily',
  'fontSize',
  'numberFormat',
  'borderPreset',
  'borderStyle',
  'merge',
  'horizontalAlignment',
  'verticalAlignment',
  'textWrapping',
  'zoom',
]);

type ColorCommandId = 'textColor' | 'fillColor' | 'borderColor';
const COLOR_COMMANDS: ReadonlySet<XlsxCommandId> = new Set<ColorCommandId>([
  'textColor',
  'fillColor',
  'borderColor',
]);
const COLOR_DEFAULTS: Record<ColorCommandId, string> = {
  textColor: '#000000',
  fillColor: '#ffffff',
  borderColor: '#000000',
};

/**
 * Arguments a control binds; required when the command takes arguments.
 * @experimental
 */
export type ToolbarCommandArgs<K extends XlsxCommandId | XlsxPluginCommandId> =
  K extends XlsxCommandId
    ? null extends XlsxCommandArgs[K]
      ? { args?: XlsxCommandArgs[K] }
      : { args: XlsxCommandArgs[K] }
    : { args?: null };

/**
 * A built-in or contributed command; a contributed one renders nothing while inactive.
 * @experimental
 */
export type ToolbarCommandButtonProps<K extends XlsxCommandId | XlsxPluginCommandId> = {
  id: K;
  /** Button content; defaults to the command's icon, or its label when it has none. */
  children?: ReactNode;
  /** Accessible name; defaults to the localized or contributed command label. */
  label?: string;
  className?: string;
} & ToolbarCommandArgs<K>;

/** @experimental */
export interface ToolbarCommandSelectProps<K extends XlsxSelectCommandId> {
  id: K;
  className?: string;
}

/** @experimental */
export interface ToolbarCommandProps<K extends XlsxCommandId | XlsxPluginCommandId> {
  id: K;
  /** Binds arguments; omit to render the command's full built-in control. */
  args?: K extends XlsxCommandId ? XlsxCommandArgs[K] : null;
  className?: string;
}

function reasonOf(state: XlsxCommandState): string | undefined {
  return state.enabled ? undefined : state.disabledReason.message;
}

function run<K extends XlsxCommandId>(
  store: XlsxCommandStore,
  id: K,
  args: XlsxCommandArgs[K],
  after?: (result: XlsxCommandResult) => void
): void {
  void store.execute(id, args).then((result) => after?.(result));
}

function focusEditor(store: XlsxCommandStore): void {
  requestAnimationFrame(() => xlsxCommandController(store)?.focusEditor());
}

function prepare<K extends XlsxCommandId>(store: XlsxCommandStore, id: K): XlsxPendingCommand<K> {
  return (
    xlsxCommandController(store)?.prepare(id) ?? {
      execute: (args: XlsxCommandArgs[K]) => store.execute(id, args),
    }
  );
}

function iconFor(id: XlsxCommandId, args: unknown, t: TFunction): ReactNode {
  const bound = args as Record<string, unknown> | null | undefined;
  if (id === 'numberFormat' && bound?.value === 'currency') {
    return <span style={{ fontSize: 16 }}>{t('toolbar.currencySymbol')}</span>;
  }
  if (id === 'numberFormat' && bound?.value === 'percent') {
    return <span style={{ fontSize: 15 }}>%</span>;
  }
  if (id === 'decimalPlaces') {
    return (
      <ToolbarIcon name={bound?.direction === 'increase' ? 'decimalIncrease' : 'decimalDecrease'} />
    );
  }
  if (id === 'fontSizeStep') {
    return <ToolbarIcon name={bound?.direction === 'increase' ? 'add' : 'remove'} />;
  }
  if (id === 'borderPreset' && bound) return <BorderGlyph preset={bound.value as BorderPreset} />;
  if (id === 'borderStyle' && bound) return <BorderStyleGlyph value={bound.value as BorderStyle} />;
  if (id === 'horizontalAlignment' && bound) {
    return <HorizontalAlignmentGlyph value={bound.value as HorizontalAlignment} />;
  }
  if (id === 'verticalAlignment' && bound) {
    return <VerticalAlignmentGlyph value={bound.value as VerticalAlignment} />;
  }
  const icon = COMMAND_ICONS[id];
  return icon ? <ToolbarIcon name={icon} /> : null;
}

function menuIconFor(id: XlsxCommandId, args: unknown, t: TFunction): ReactNode {
  if (id === 'numberFormat' || id === 'merge') return undefined;
  return iconFor(id, args, t) ?? undefined;
}

function testIdFor(id: XlsxCommandId, args: unknown): string | undefined {
  if (id === 'merge' && (args as { value?: unknown } | undefined)?.value === 'all') {
    return 'xlsx-merge-all';
  }
  return args === undefined || args === null ? TEST_IDS[id] : undefined;
}

/** What overflow presentations need beyond the store: value prompts and a color picker. */
export type OverflowTools = Pick<OverflowRegistry, 'prompt' | 'pickColor'> | null;

function commandItem<K extends XlsxCommandId>(
  store: XlsxCommandStore,
  t: TFunction,
  id: K,
  args: XlsxCommandArgs[K] | undefined,
  label?: string
): OverflowMenuEntry {
  const state = store.getState(id, args) as XlsxCommandState;
  const toggle = TOGGLES.has(id) && state.active !== undefined;
  return {
    kind: 'item',
    id: args === undefined ? id : `${id}:${JSON.stringify(args)}`,
    label: label ?? t(commandLabelKey(id, args)),
    icon: menuIconFor(id, args, t),
    shortcut: commandShortcut(id, args) ?? undefined,
    checked: toggle ? state.active : undefined,
    disabled: !state.enabled,
    description: reasonOf(state),
    onSelect: () => run(store, id, (args ?? null) as XlsxCommandArgs[K]),
  };
}

function optionItems<K extends XlsxSelectCommandId>(
  store: XlsxCommandStore,
  t: TFunction,
  id: K,
  state: XlsxCommandState<K> = store.getState(id)
): OverflowMenuEntry[] {
  const radio = id !== 'merge';
  return (state.options ?? []).flatMap((option, index): OverflowMenuEntry[] => {
    const optionState = option.state as XlsxCommandState;
    const item: OverflowMenuEntry = {
      kind: 'item',
      id: `${id}-${index}`,
      label: option.label,
      icon: menuIconFor(id, option.args, t),
      ...(radio ? { radio: true, checked: optionState.active === true } : {}),
      disabled: !optionState.enabled,
      description: reasonOf(optionState),
      onSelect: () => run(store, id, option.args),
    };
    const unmerge = id === 'merge' && (option.args as { value?: string }).value === 'unmerge';
    return unmerge ? [{ kind: 'separator', id: `${id}-separator` }, item] : [item];
  });
}

function colorItem(
  store: XlsxCommandStore,
  t: TFunction,
  id: ColorCommandId,
  tools: OverflowTools
): OverflowMenuEntry {
  const state = store.getState(id);
  const value = isHexColor(state.value) ? state.value : COLOR_DEFAULTS[id];
  return {
    kind: 'item',
    id,
    label: t(commandLabelKey(id)),
    icon: <ColorSwatch color={value} />,
    disabled: !state.enabled || !tools,
    description: reasonOf(state as XlsxCommandState),
    onSelect: () => {
      const pending = prepare(store, id);
      tools?.pickColor({ value, submit: (color) => void pending.execute({ color }) });
    },
  };
}

function promptItem(
  store: XlsxCommandStore,
  t: TFunction,
  id: 'fontSize' | 'zoom',
  tools: OverflowTools
): OverflowMenuEntry | null {
  if (!tools) return null;
  const state = store.getState(id) as XlsxCommandState;
  const zoom = id === 'zoom';
  return {
    kind: 'item',
    id: `${id}-custom`,
    label: t(zoom ? 'commands.customZoom' : 'commands.customFontSize'),
    disabled: !state.enabled,
    description: reasonOf(state),
    onSelect: () => {
      const pending = prepare(store, id);
      const current = typeof state.value === 'number' ? state.value : null;
      tools.prompt({
        title: t(commandLabelKey(id)),
        label: t(zoom ? 'commands.zoomPercent' : 'commands.fontSizePoints'),
        initialValue: current === null ? '' : String(zoom ? Math.round(current * 100) : current),
        valid: (value) => {
          const number = Number(value);
          if (!/^\d+(\.\d+)?$/.test(value)) return false;
          return zoom
            ? number / 100 >= MIN_ZOOM && number / 100 <= MAX_ZOOM
            : number >= MIN_FONT_POINTS && number <= MAX_FONT_POINTS;
        },
        submit: (value) =>
          void (zoom
            ? (pending as { execute(args: XlsxCommandArgs['zoom']): Promise<unknown> }).execute({
                scale: Number(value) / 100,
              })
            : (pending as { execute(args: XlsxCommandArgs['fontSize']): Promise<unknown> }).execute({
                points: Number(value),
              })),
      });
    },
  };
}

function borderEntries(
  store: XlsxCommandStore,
  t: TFunction,
  tools: OverflowTools
): OverflowMenuEntry[] {
  return [
    ...optionItems(store, t, 'borderPreset'),
    { kind: 'separator', id: 'borders-styles' },
    ...optionItems(store, t, 'borderStyle'),
    { kind: 'separator', id: 'borders-color' },
    colorItem(store, t, 'borderColor', tools),
  ];
}

/** The overflow-menu presentation of one command. */
export function commandOverflowEntry<K extends XlsxCommandId>(
  store: XlsxCommandStore,
  t: TFunction,
  id: K,
  args?: XlsxCommandArgs[K],
  tools: OverflowTools = null,
  label?: string
): OverflowMenuEntry {
  if (args === undefined) {
    if (COLOR_COMMANDS.has(id)) return colorItem(store, t, id as ColorCommandId, tools);
    if (SELECT_COMMANDS.has(id)) {
      const state = store.getState(id) as XlsxCommandState;
      const entries =
        id === 'borderPreset'
          ? borderEntries(store, t, tools)
          : optionItems(store, t, id as XlsxSelectCommandId);
      const custom =
        id === 'fontSize' || id === 'zoom' ? promptItem(store, t, id, tools) : null;
      return {
        kind: 'submenu',
        id,
        label: label ?? t(commandLabelKey(id)),
        disabled: !state.enabled,
        description: reasonOf(state),
        entries: custom ? [...entries, custom] : entries,
      };
    }
  }
  return commandItem(store, t, id, args, label);
}

function useCommandOverflow<K extends XlsxCommandId>(
  element: RefObject<HTMLElement | null>,
  id: K,
  args?: XlsxCommandArgs[K]
): void {
  const store = useXlsxCommands();
  const { t } = useTranslation();
  const tools = useContext(ToolbarOverflowContext);
  useOverflowSource(element, () => [commandOverflowEntry(store, t, id, args, tools)]);
}

/**
 * A button bound to one command, showing its pressed and disabled state.
 * @experimental
 */
export function ToolbarCommandButton<K extends XlsxCommandId | XlsxPluginCommandId>(
  props: ToolbarCommandButtonProps<K>
) {
  if (isPluginCommandId(props.id)) {
    return (
      <PluginCommandButton
        id={props.id}
        label={props.label}
        className={props.className}
      >
        {props.children}
      </PluginCommandButton>
    );
  }
  const Button = BuiltInCommandButton as ComponentType<ToolbarCommandButtonProps<XlsxCommandId>>;
  return <Button {...(props as ToolbarCommandButtonProps<XlsxCommandId>)} />;
}

function PluginCommandButton({
  id,
  label,
  className,
  children,
}: {
  id: XlsxPluginCommandId;
  label?: string;
  className?: string;
  children?: ReactNode;
}) {
  const command = useXlsxCommand(id);
  if (command.descriptor === null) return null;
  const state = command.state;
  const name = label ?? command.label;
  const toggle = state.active !== undefined;
  return (
    <ToolbarButtonBase
      active={toggle ? state.active : undefined}
      toggle={toggle}
      disabled={!state.enabled}
      description={state.enabled ? undefined : state.disabledReason.message}
      title={name}
      shortcut={command.shortcut}
      className={className}
      style={children ? undefined : { padding: '0 8px' }}
      onClick={() => void command.execute()}
    >
      {children ?? name}
    </ToolbarButtonBase>
  );
}

function BuiltInCommandButton<K extends XlsxCommandId>(props: ToolbarCommandButtonProps<K>) {
  const { id, children, label, className } = props;
  const args = (props as { args?: XlsxCommandArgs[K] }).args;
  const store = useXlsxCommands();
  const command = useXlsxCommand(id, args);
  const { t } = useTranslation();
  const tools = useContext(ToolbarOverflowContext);
  const state = command.state as XlsxCommandState;
  const name = label ?? command.label;
  const toggle = TOGGLES.has(id) && state.active !== undefined;
  return (
    <ToolbarButtonBase
      active={toggle ? state.active : undefined}
      toggle={toggle}
      disabled={!state.enabled}
      description={reasonOf(state)}
      title={name}
      shortcut={command.shortcut}
      testId={testIdFor(id, args)}
      className={className}
      onClick={() => void command.execute()}
      overflow={() => [commandOverflowEntry(store, t, id, args, tools, name)]}
    >
      {children ?? iconFor(id, args, t) ?? name}
    </ToolbarButtonBase>
  );
}

function OptionMenu<K extends XlsxSelectCommandId>({
  id,
  close,
}: {
  id: K;
  close: () => void;
}) {
  const store = useXlsxCommands();
  const { t } = useTranslation();
  const state = useXlsxCommandState(id) as XlsxCommandState<K>;
  return (
    <>
      {optionItems(store, t, id, state).map((entry) =>
        entry.kind === 'item' ? (
          <ToolbarMenuItem
            key={entry.id}
            label={entry.label}
            icon={entry.icon}
            selected={entry.radio ? entry.checked === true : undefined}
            disabled={entry.disabled}
            description={entry.description}
            onClick={entry.onSelect}
            close={close}
          />
        ) : (
          <ToolbarMenuSeparator key={entry.id} />
        )
      )}
    </>
  );
}

function HiddenColorInput({
  inputRef,
  onPick,
}: {
  inputRef: RefObject<HTMLInputElement | null>;
  onPick(color: string): void;
}) {
  return (
    <input
      ref={inputRef}
      type="color"
      tabIndex={-1}
      aria-hidden="true"
      onChange={(event) => onPick(event.target.value)}
      style={{
        position: 'absolute',
        width: 1,
        height: 1,
        padding: 0,
        border: 0,
        opacity: 0,
        pointerEvents: 'none',
      }}
    />
  );
}

/** Opens the platform color picker of `input` at `value`. */
export function openColorPicker(input: HTMLInputElement | null, value: string): void {
  if (!input) return;
  input.value = value;
  const picker = input as HTMLInputElement & { showPicker?: () => void };
  try {
    if (picker.showPicker) picker.showPicker();
    else input.click();
  } catch {
    input.click();
  }
}

function BordersControl() {
  const store = useXlsxCommands();
  const { t } = useTranslation();
  const state = useXlsxCommandState('borderPreset');
  const color = useXlsxCommandState('borderColor');
  const inputRef = useRef<HTMLInputElement>(null);
  const pending = useRef<XlsxPendingCommand<'borderColor'> | null>(null);
  const current = isHexColor(color.value) ? color.value : COLOR_DEFAULTS.borderColor;
  return (
    <>
      <ToolbarDropdown
        title={t('toolbar.borders')}
        disabled={!state.enabled}
        description={reasonOf(state as XlsxCommandState)}
        menuWidth={240}
        trigger={
          <>
            <BorderGlyph preset={state.value ?? 'all'} />
            <ToolbarIcon name="chevronDown" size={12} />
          </>
        }
      >
        {(close) => (
          <>
            <OptionMenu id="borderPreset" close={close} />
            <ToolbarMenuSeparator />
            <OptionMenu id="borderStyle" close={close} />
            <ToolbarMenuSeparator />
            <ToolbarMenuItem
              label={t('toolbar.borderColor')}
              icon={<ColorSwatch color={current} />}
              disabled={!color.enabled}
              description={reasonOf(color as XlsxCommandState)}
              onClick={() => {
                pending.current = prepare(store, 'borderColor');
                openColorPicker(inputRef.current, current);
              }}
              close={close}
            />
          </>
        )}
      </ToolbarDropdown>
      <HiddenColorInput
        inputRef={inputRef}
        onPick={(picked) => void pending.current?.execute({ color: picked })}
      />
    </>
  );
}

function MergeControl() {
  const store = useXlsxCommands();
  const { t } = useTranslation();
  const state = useXlsxCommandState('merge');
  const all = useXlsxCommandState('merge', { value: 'all' });
  return (
    <>
      <ToolbarButtonBase
        title={t('toolbar.merge.all')}
        disabled={!all.enabled}
        description={reasonOf(all as XlsxCommandState)}
        onClick={() => run(store, 'merge', { value: 'all' })}
        style={{ borderRadius: '4px 0 0 4px' }}
        testId="xlsx-merge-all"
        overflow={null}
      >
        <ToolbarIcon name="merge" />
      </ToolbarButtonBase>
      <ToolbarDropdown
        title={t('toolbar.mergeCells')}
        disabled={!state.enabled}
        description={reasonOf(state as XlsxCommandState)}
        menuWidth={220}
        style={{ minWidth: 20, width: 20, padding: 0, borderRadius: '0 4px 4px 0' }}
        trigger={<ToolbarIcon name="chevronDown" size={13} />}
      >
        {(close) => <OptionMenu id="merge" close={close} />}
      </ToolbarDropdown>
    </>
  );
}

const DROPDOWN_WIDTHS: Partial<Record<XlsxSelectCommandId, number>> = {
  fontFamily: 210,
  numberFormat: 220,
  borderStyle: 220,
};

function dropdownTrigger(
  id: XlsxSelectCommandId,
  state: XlsxCommandState,
  t: TFunction
): ReactNode {
  const chevron = <ToolbarIcon name="chevronDown" size={id === 'numberFormat' || id === 'fontFamily' ? 13 : 12} />;
  switch (id) {
    case 'fontFamily':
      return (
        <>
          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis' }}>
            {typeof state.value === 'string' ? state.value : ''}
          </span>
          {chevron}
        </>
      );
    case 'numberFormat':
      return (
        <>
          <span>123</span>
          {chevron}
        </>
      );
    case 'horizontalAlignment':
      return (
        <>
          <HorizontalAlignmentGlyph value={(state.value as HorizontalAlignment | null) ?? 'left'} />
          {chevron}
        </>
      );
    case 'verticalAlignment':
      return (
        <>
          <VerticalAlignmentGlyph value={(state.value as VerticalAlignment | null) ?? 'middle'} />
          {chevron}
        </>
      );
    case 'textWrapping':
      return (
        <>
          <ToolbarIcon name="wrap" />
          {chevron}
        </>
      );
    case 'borderStyle':
      return (
        <>
          <BorderStyleGlyph value={(state.value as BorderStyle | null) ?? 'solid'} />
          {chevron}
        </>
      );
    default:
      return <span>{t(commandLabelKey(id))}</span>;
  }
}

/**
 * The built-in picker of a selector command.
 * @experimental
 */
export function ToolbarCommandSelect<K extends XlsxSelectCommandId>({
  id,
  className,
}: ToolbarCommandSelectProps<K>) {
  const store = useXlsxCommands();
  const { t } = useTranslation();
  const state = useXlsxCommandState(id) as XlsxCommandState;
  const wrapperRef = useRef<HTMLSpanElement>(null);
  useCommandOverflow(wrapperRef, id);
  const disabled = !state.enabled;
  const description = reasonOf(state);

  let control: ReactNode;
  if (id === 'fontSize') {
    const value = typeof state.value === 'number' ? String(state.value) : '';
    control = (
      <EditableCombobox
        value={value}
        options={(state.options ?? []).map((option) => ({
          value: String((option.args as XlsxCommandArgs['fontSize']).points),
          label: option.label,
        }))}
        label={t('toolbar.fontSize')}
        disabled={disabled}
        description={description}
        onCommit={(input, pointer) => {
          const points = Number.parseFloat(input);
          if (!Number.isFinite(points) || points < MIN_FONT_POINTS || points > MAX_FONT_POINTS) {
            return;
          }
          run(store, 'fontSize', { points }, pointer ? () => focusEditor(store) : undefined);
        }}
        width={50}
        inputStyle={{ textAlign: 'center' }}
      />
    );
  } else if (id === 'zoom') {
    const percent = `${Math.round((typeof state.value === 'number' ? state.value : 1) * 100)}%`;
    control = (
      <EditableCombobox
        value={percent}
        options={(state.options ?? []).map((option) => ({
          value: option.label,
          label: option.label,
        }))}
        label={t('toolbar.zoomValue', { value: percent })}
        disabled={disabled}
        description={description}
        onCommit={(input) => {
          const scale = Number.parseFloat(input.replace('%', '')) / 100;
          if (Number.isFinite(scale) && scale >= MIN_ZOOM && scale <= MAX_ZOOM) {
            run(store, 'zoom', { scale });
          }
        }}
        width={76}
        testId="xlsx-zoom"
      />
    );
  } else if (id === 'borderPreset') {
    control = <BordersControl />;
  } else if (id === 'merge') {
    control = <MergeControl />;
  } else {
    control = (
      <ToolbarDropdown
        title={t(
          id === 'numberFormat' ? 'toolbar.moreNumberFormats' : commandLabelKey(id)
        )}
        disabled={disabled}
        description={description}
        menuWidth={DROPDOWN_WIDTHS[id] ?? 180}
        style={id === 'fontFamily' ? { width: 120, justifyContent: 'space-between' } : undefined}
        trigger={dropdownTrigger(id, state, t)}
      >
        {(close) => <OptionMenu id={id} close={close} />}
      </ToolbarDropdown>
    );
  }
  return (
    <span
      ref={wrapperRef}
      className={className}
      style={{ display: 'inline-flex', alignItems: 'center', flex: '0 0 auto' }}
    >
      {control}
    </span>
  );
}

function CommandColorPicker({ id, className }: { id: ColorCommandId; className?: string }) {
  const store = useXlsxCommands();
  const { t } = useTranslation();
  const state = useXlsxCommandState(id);
  const wrapperRef = useRef<HTMLSpanElement>(null);
  const pending = useRef<XlsxPendingCommand<ColorCommandId> | null>(null);
  useCommandOverflow(wrapperRef, id);
  return (
    <span ref={wrapperRef} className={className} style={{ display: 'inline-flex', flex: '0 0 auto' }}>
      <ColorPicker
        mode={id === 'textColor' ? 'text' : id === 'fillColor' ? 'fill' : 'border'}
        value={isHexColor(state.value) ? state.value : COLOR_DEFAULTS[id]}
        label={t(commandLabelKey(id))}
        disabled={!state.enabled}
        description={reasonOf(state as XlsxCommandState)}
        onOpen={() => {
          pending.current = prepare(store, id);
        }}
        onChange={(color) => void (pending.current ?? prepare(store, id)).execute({ color })}
      />
    </span>
  );
}

/**
 * The built-in control of any command, as the default toolbar presents it.
 * @experimental
 */
export function ToolbarCommand<K extends XlsxCommandId | XlsxPluginCommandId>(
  props: ToolbarCommandProps<K>
) {
  const { id, args, className } = props;
  if (isPluginCommandId(id)) {
    const contributed: XlsxPluginCommandId = id;
    return <ToolbarCommandButton id={contributed} className={className} />;
  }
  if (args === undefined) {
    if (COLOR_COMMANDS.has(id)) {
      return <CommandColorPicker id={id as ColorCommandId} className={className} />;
    }
    if (SELECT_COMMANDS.has(id)) {
      return <ToolbarCommandSelect id={id as XlsxSelectCommandId} className={className} />;
    }
  }
  const CommandButton = ToolbarCommandButton as unknown as ComponentType<{
    id: XlsxCommandId;
    args?: unknown;
    className?: string;
  }>;
  return <CommandButton id={id} args={args} className={className} />;
}
