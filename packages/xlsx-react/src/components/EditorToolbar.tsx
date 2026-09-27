import { useContext, useLayoutEffect, useMemo, useRef } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { Translations } from '@betteroffice/xlsx-i18n';
import { xlsxCommandController } from '../commands/createXlsxCommandStore';
import { useXlsxChrome, useXlsxCommands, useXlsxCommandState } from '../commands/hooks';
import { LocaleProvider } from '../i18n';
import {
  EditorChromeContext,
  EditorToolbarContext,
  ToolbarModeContext,
  type EditorToolbarProps,
} from './EditorToolbarContext';
import { FormulaBar } from './toolbar/FormulaBar';
import { commandForAction } from './toolbar/legacyToolbarStore';
import { assertCommandProps, Toolbar, type ToolbarProps } from './Toolbar';

interface EditorToolbarComponent {
  (props: EditorToolbarProps): React.JSX.Element;
  Toolbar: typeof Toolbar;
  FormulaBar: typeof FormulaBar;
}

function Frame({
  className,
  style,
  children,
}: {
  className?: string;
  style?: CSSProperties;
  children: ReactNode;
}) {
  return (
    <div
      className={className}
      data-testid="xlsx-editor-toolbar"
      style={{
        display: 'flex',
        flexDirection: 'column',
        flex: '0 0 auto',
        ...style,
      }}
    >
      {children}
    </div>
  );
}

/** Command state projected onto the legacy props, so `useEditorToolbar` keeps working. */
function LegacyProjection({ children }: { children: ReactNode }) {
  const store = useXlsxCommands();
  const paintFormat = useXlsxCommandState('paintFormat');
  const numberFormat = useXlsxCommandState('numberFormat');
  const fontFamily = useXlsxCommandState('fontFamily');
  const fontSize = useXlsxCommandState('fontSize');
  const bold = useXlsxCommandState('bold');
  const italic = useXlsxCommandState('italic');
  const strikethrough = useXlsxCommandState('strikethrough');
  const textColor = useXlsxCommandState('textColor');
  const fillColor = useXlsxCommandState('fillColor');
  const borderPreset = useXlsxCommandState('borderPreset');
  const borderStyle = useXlsxCommandState('borderStyle');
  const borderColor = useXlsxCommandState('borderColor');
  const horizontalAlignment = useXlsxCommandState('horizontalAlignment');
  const verticalAlignment = useXlsxCommandState('verticalAlignment');
  const textWrapping = useXlsxCommandState('textWrapping');
  const merge = useXlsxCommandState('merge');
  const unmerge = useXlsxCommandState('merge', { value: 'unmerge' });
  const undo = useXlsxCommandState('undo');
  const redo = useXlsxCommandState('redo');
  const zoom = useXlsxCommandState('zoom');
  const value = useMemo<ToolbarProps>(() => {
    const mark = (active: boolean | 'mixed' | undefined) => (active === 'mixed' ? undefined : active);
    const defined = <T,>(value: T | null | undefined) => value ?? undefined;
    return {
      currentFormatting: {
        paintFormat: paintFormat.active === true,
        numberFormat: defined(numberFormat.value?.kind),
        numberFormatPattern: defined(numberFormat.value?.pattern),
        fontFamily: defined(fontFamily.value),
        fontSize: defined(fontSize.value),
        bold: mark(bold.active),
        italic: mark(italic.active),
        strikethrough: mark(strikethrough.active),
        textColor: defined(textColor.value),
        fillColor: defined(fillColor.value),
        borderPreset: defined(borderPreset.value),
        borderStyle: defined(borderStyle.value),
        borderColor: defined(borderColor.value),
        horizontalAlignment: defined(horizontalAlignment.value),
        verticalAlignment: defined(verticalAlignment.value),
        textWrapping: defined(textWrapping.value),
      },
      selectionShape: merge.value
        ? { rows: merge.value.rows, columns: merge.value.columns, canUnmerge: unmerge.enabled }
        : undefined,
      onFormat: (action) => {
        const call = commandForAction(action);
        void store.execute(call.id, call.args as never);
      },
      onMerge: (action) => void store.execute('merge', { value: action }),
      onSearchMenus: () => void store.execute('searchMenus', null),
      onUndo: () => void store.execute('undo', null),
      onRedo: () => void store.execute('redo', null),
      canUndo: undo.enabled,
      canRedo: redo.enabled,
      onPrint: () => void store.execute('print', null),
      zoom: zoom.value,
      onZoomChange: (scale) => void store.execute('zoom', { scale }),
    };
  }, [
    store,
    paintFormat,
    numberFormat,
    fontFamily,
    fontSize,
    bold,
    italic,
    strikethrough,
    textColor,
    fillColor,
    borderPreset,
    borderStyle,
    borderColor,
    horizontalAlignment,
    verticalAlignment,
    textWrapping,
    merge,
    unmerge,
    undo,
    redo,
    zoom,
  ]);
  return <EditorToolbarContext.Provider value={value}>{children}</EditorToolbarContext.Provider>;
}

/** Locale and shortcut ownership for command chrome rendered outside the editor. */
function ExternalChrome({
  i18n,
  children,
}: {
  i18n: Translations | undefined;
  children: ReactNode;
}) {
  const controller = xlsxCommandController(useXlsxCommands());
  const chrome = useXlsxChrome();
  const rootRef = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const root = rootRef.current;
    if (!root || !controller) return;
    return controller.registerChrome(root);
  }, [controller]);
  return (
    <div ref={rootRef} style={{ display: 'contents' }}>
      <LocaleProvider i18n={i18n ?? chrome?.i18n}>{children}</LocaleProvider>
    </div>
  );
}

function EditorToolbarBase({ children, className, style, i18n, mode, ...toolbarProps }: EditorToolbarProps) {
  const insideEditor = useContext(EditorChromeContext);
  if (mode !== 'commands') {
    return (
      <ToolbarModeContext.Provider value="legacy">
        <EditorToolbarContext.Provider value={toolbarProps}>
          <Frame className={className} style={style}>
            {children}
          </Frame>
        </EditorToolbarContext.Provider>
      </ToolbarModeContext.Provider>
    );
  }
  assertCommandProps(toolbarProps, 'EditorToolbar');
  const content = (
    <ToolbarModeContext.Provider value="commands">
      <LegacyProjection>
        <Frame className={className} style={style}>
          {children ?? <Toolbar />}
        </Frame>
      </LegacyProjection>
    </ToolbarModeContext.Provider>
  );
  return insideEditor ? content : <ExternalChrome i18n={i18n}>{content}</ExternalChrome>;
}

/**
 * The editor chrome as a compound component. `mode="commands"` binds its
 * parts to the nearest editor's commands (inside `XlsxEditor` or an
 * `XlsxCommandProvider`) and, without children, renders the formatting rail.
 * The default legacy mode shares its props with the parts, as before.
 */
const EditorToolbar = EditorToolbarBase as EditorToolbarComponent;
EditorToolbar.Toolbar = Toolbar;
EditorToolbar.FormulaBar = FormulaBar;

export { EditorToolbar };
