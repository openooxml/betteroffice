import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore, type KeyboardEvent, type ReactNode } from 'react';
import { defineDocxPlugin } from '../defineDocxPlugin';
import { useDocxSelectionPlacement } from '../selectionPlacement';
import type { DocxPlugin, DocxPluginContext, DocxPluginGeometry, DocxPluginRect } from '../types';
import { useTranslation } from '../../i18n/LocaleContext';
import { IconBold, IconFormatClear, IconItalic, IconStrikethrough, IconUnderline } from '../../components/ui/Icons';
import {
  INITIAL_MINI_TOOLBAR_STATE,
  MINI_TOOLBAR_GAP,
  MINI_TOOLBAR_PADDING,
  afterSelectionChange,
  avoidCommentButton,
  buttonState,
  dismissMiniToolbar,
  miniToolbarEligible,
  miniToolbarSize,
  miniToolbarView,
  resolveMiniToolbarItems,
  rovingIndex,
  type DocxMiniToolbarItem,
  type MiniToolbarState,
} from './model';

/** @experimental The plugin API may change in minor releases. */
export const DOCX_MINI_TOOLBAR_ID = 'betteroffice.mini-toolbar';

/** @experimental The plugin API may change in minor releases. */
export interface DocxMiniToolbarOptions {
  /** The buttons in display order; all six by default. An unknown or repeated id throws a `TypeError`. */
  items?: readonly DocxMiniToolbarItem[];
}

type Context = DocxPluginContext<MiniToolbarState>;

const ICON_SIZE = 18;

function PasteIcon({ size }: { size: number }) {
  return (
    <svg xmlns="http://www.w3.org/2000/svg" width={size} height={size} viewBox="0 -960 960 960" fill="currentColor" aria-hidden="true">
      <path d="M200-120q-33 0-56.5-23.5T120-200v-560q0-33 23.5-56.5T200-840h167q11-35 43-57.5t70-22.5q40 0 71.5 22.5T594-840h166q33 0 56.5 23.5T840-760v560q0 33-23.5 56.5T760-120H200Zm0-80h560v-560h-80v120H280v-120h-80v560Zm280-560q17 0 28.5-11.5T520-800q0-17-11.5-28.5T480-840q-17 0-28.5 11.5T440-800q0 17 11.5 28.5T480-760Z" />
    </svg>
  );
}

const ICONS: Record<DocxMiniToolbarItem, (props: { size: number }) => ReactNode> = {
  bold: IconBold,
  italic: IconItalic,
  underline: IconUnderline,
  strikethrough: IconStrikethrough,
  clearFormatting: IconFormatClear,
  pasteAsPlainText: PasteIcon,
};

const COARSE = '(pointer: coarse)';
const subscribeCoarse = (listener: () => void) => {
  const query = globalThis.matchMedia?.(COARSE);
  query?.addEventListener?.('change', listener);
  return () => query?.removeEventListener?.('change', listener);
};
const readCoarse = () => globalThis.matchMedia?.(COARSE)?.matches ?? false;

const eligible = (context: Context) => miniToolbarEligible(context.snapshot, context.state);
const toolbarActive = (context: Context) => context.state.shown && eligible(context);
const FOCUS_COMMAND = 'focus';
const NOT_SHOWN = Object.freeze({ code: 'not-shown', message: 'The mini toolbar is not shown' });

function MiniToolbar({ context, geometry, items }: { context: Context; geometry: DocxPluginGeometry; items: readonly DocxMiniToolbarItem[] }) {
  const { t } = useTranslation();
  const coarse = useSyncExternalStore(subscribeCoarse, readCoarse, () => false);
  const size = miniToolbarSize(items.length, coarse);
  const placement = useDocxSelectionPlacement(context, { width: size.width, height: size.height, geometry });
  const page = useRef<DocxPluginRect | null>(null);
  const answer = placement ? geometry.getAnchorGeometry({ kind: 'selection' }) : null;
  if (answer?.ok) page.current = answer.pageRect;
  const toolbar = useRef<HTMLDivElement | null>(null);
  const latest = useRef(context);
  latest.current = context;
  const holdsFocus = toolbar.current?.contains((toolbar.current.getRootNode() as Document | ShadowRoot).activeElement) === true;
  const toolbarFocused = context.state.toolbarFocused || holdsFocus;
  const view = miniToolbarView(miniToolbarEligible(context.snapshot, { ...context.state, toolbarFocused }), placement);
  const shown = view !== 'hidden';
  const commandStates = () => items.map((item) => JSON.stringify(context.commands.getState(item))).join('\n');
  useSyncExternalStore(context.commands.subscribe, commandStates, commandStates);
  const states = items.map((item) => buttonState(item, context.commands.getState(item)));
  const [current, setCurrent] = useState(0);
  const firstEnabled = states.findIndex((state) => !state.disabled);
  const tabStop = states[current] && !states[current]!.disabled ? current : Math.max(0, firstEnabled);
  const buttons = useRef<(HTMLButtonElement | null)[]>([]);
  const handled = useRef(context.state.focusRequest);

  useEffect(() => {
    if (context.state.shown === shown) return;
    void context.run((next) => { next.setState((state) => ({ ...state, shown })); });
  }, [context, shown]);
  useEffect(() => {
    const element = toolbar.current;
    if (!shown || !element) return;
    const owner = element.ownerDocument;
    const onPress = (event: PointerEvent) => {
      if (element.contains(event.target as Node | null)) return;
      void latest.current.run((next) => { next.setState((state) => dismissMiniToolbar(state, next.snapshot)); });
    };
    owner.addEventListener('pointerdown', onPress, true);
    return () => owner.removeEventListener('pointerdown', onPress, true);
  }, [shown]);
  useEffect(() => {
    if (context.state.focusRequest === handled.current) return;
    handled.current = context.state.focusRequest;
    if (firstEnabled < 0) return;
    setCurrent(firstEnabled);
    buttons.current[firstEnabled]?.focus();
  });
  useLayoutEffect(() => {
    if (toolbar.current) toolbar.current.inert = view === 'kept';
  });

  if (!placement || view === 'hidden') return null;

  const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape' || event.key === 'Tab') {
      event.preventDefault();
      const dismiss = event.key === 'Escape';
      void context.run((next) => {
        next.focusDocument();
        if (dismiss) next.setState((state) => dismissMiniToolbar(state, next.snapshot));
      });
      return;
    }
    const from = buttons.current.findIndex((button) => button === event.target);
    const to = rovingIndex(event.key, from < 0 ? tabStop : from, states.map((state) => !state.disabled));
    if (to === null) return;
    event.preventDefault();
    setCurrent(to);
    buttons.current[to]?.focus();
  };

  return (
    <div
      ref={toolbar}
      role="toolbar"
      aria-label={t('plugins.miniToolbar')}
      aria-orientation="horizontal"
      data-docx-mini-toolbar={view}
      onMouseDown={(event) => event.preventDefault()}
      onFocus={() => {
        if (!context.state.toolbarFocused) void context.run((next) => { next.setState((state) => ({ ...state, toolbarFocused: true })); });
      }}
      onBlur={(event) => {
        if (event.currentTarget.contains(event.relatedTarget as Node | null)) return;
        void context.run((next) => { next.setState((state) => ({ ...state, toolbarFocused: false })); });
      }}
      onKeyDown={onKeyDown}
      style={{
        position: 'absolute',
        left: avoidCommentButton(placement.x, size.width, page.current),
        top: placement.y,
        width: size.width,
        height: size.height,
        boxSizing: 'border-box',
        display: 'flex',
        alignItems: 'center',
        gap: MINI_TOOLBAR_GAP,
        padding: MINI_TOOLBAR_PADDING,
        borderRadius: 8,
        border: '1px solid var(--doc-border)',
        backgroundColor: 'var(--doc-surface)',
        boxShadow: '0 2px 8px var(--doc-shadow)',
        opacity: view === 'kept' ? 0.5 : 1,
        pointerEvents: view === 'kept' ? 'none' : 'auto',
      }}
    >
      {items.map((item, index) => {
        const Icon = ICONS[item];
        const state = states[index]!;
        const label = t(context.commands.getDescriptor(item).labelKey);
        return (
          <button
            key={item}
            ref={(element) => { buttons.current[index] = element; }}
            type="button"
            className="oox-toolbar-toggle inline-flex items-center justify-center rounded"
            aria-label={label}
            title={label}
            aria-pressed={state.pressed}
            data-active={state.pressed === 'true' ? 'true' : undefined}
            disabled={state.disabled}
            tabIndex={index === tabStop ? 0 : -1}
            onFocus={() => setCurrent(index)}
            onClick={() => void context.run(async (next) => { await next.commands.execute(item, null); })}
            style={{ width: size.button, height: size.button, flexShrink: 0, border: 0, background: 'transparent', opacity: state.disabled ? 0.38 : 1 }}
          >
            <Icon size={ICON_SIZE} />
          </button>
        );
      })}
    </div>
  );
}

/**
 * A floating toolbar over a settled text selection in the document body while the document input or
 * the toolbar has focus, with built-in formatting commands. A pointer press outside it hides it until
 * the selection changes. Off unless the host adds it to `plugins`; grant it `document: 'write'`
 * and the ids of its items under `DOCX_MINI_TOOLBAR_ID`, or its buttons stay disabled.
 *
 * @experimental The plugin API may change in minor releases.
 */
export function docxMiniToolbar(options: DocxMiniToolbarOptions = {}): DocxPlugin {
  const items = resolveMiniToolbarItems(options.items);
  return defineDocxPlugin<MiniToolbarState>({
    id: DOCX_MINI_TOOLBAR_ID,
    revision: items.join(','),
    createState: () => INITIAL_MINI_TOOLBAR_STATE,
    onEvent(context, event) {
      if (event.type !== 'selection-change') return;
      const next = afterSelectionChange(context.state, event.selection);
      if (next !== context.state) context.setState(next);
    },
    overlay: ({ context, geometry }) => <MiniToolbar context={context} geometry={geometry} items={items} />,
    commands: [
      {
        id: FOCUS_COMMAND,
        label: 'Focus the mini toolbar',
        mutatesDocument: false,
        shortcuts: ['Alt+F10'],
        getState: (context) => (toolbarActive(context) ? { enabled: true } : { enabled: false, disabledReason: NOT_SHOWN }),
        execute(context) {
          context.setState((state) => ({ ...state, focusRequest: state.focusRequest + 1 }));
          return { ok: true, status: 'requested' };
        },
      },
    ],
  });
}
