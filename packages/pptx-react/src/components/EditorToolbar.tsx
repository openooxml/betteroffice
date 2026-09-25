import { useContext, useLayoutEffect, useRef } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type { Translations } from '@betteroffice/pptx-i18n';
import { pptxCommandController } from '../commands/createPptxCommandStore';
import { usePptxChrome, usePptxCommands } from '../commands/hooks';
import { LocaleProvider } from '../i18n';
import {
  EditorChromeContext,
  EditorToolbarContext,
  ToolbarModeContext,
} from './EditorToolbarContext';
import type { EditorToolbarProps } from './EditorToolbarContext';
import { useLegacyProjection } from './toolbar/legacyCommands';
import { rejectLegacyState, Toolbar, type CommandToolbarProps } from './Toolbar';

/** A toolbar region bound to the nearest editor's or `PptxCommandProvider`'s commands. */
export type EditorToolbarCommandProps = Omit<CommandToolbarProps, 'children'> & {
  /** Compound parts; defaults to the formatting rail with the built-in controls. */
  children?: ReactNode;
  /** Locale when rendered outside the editor; defaults to the editor's. */
  i18n?: Translations;
};

interface EditorToolbarComponent {
  (
    props:
      | (EditorToolbarProps & { children: ReactNode; style?: CSSProperties })
      | EditorToolbarCommandProps
  ): React.JSX.Element;
  Toolbar: typeof Toolbar;
}

const rootStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  flex: '0 0 auto',
};

function ExternalChrome({
  i18n,
  children,
}: {
  i18n: Translations | undefined;
  children: ReactNode;
}) {
  const controller = pptxCommandController(usePptxCommands());
  const chrome = usePptxChrome();
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

function CommandEditorToolbar(props: EditorToolbarCommandProps) {
  rejectLegacyState(props, 'EditorToolbar');
  const { children, className, style, i18n } = props;
  const projection = useLegacyProjection(usePptxCommands());
  const inEditor = useContext(EditorChromeContext);
  const toolbar = (
    <ToolbarModeContext.Provider value="commands">
      <EditorToolbarContext.Provider value={projection}>
        <div
          className={className}
          data-testid="pptx-editor-toolbar"
          style={{ ...rootStyle, ...style }}
        >
          {children ?? <Toolbar mode="commands" />}
        </div>
      </EditorToolbarContext.Provider>
    </ToolbarModeContext.Provider>
  );
  return inEditor ? toolbar : <ExternalChrome i18n={i18n}>{toolbar}</ExternalChrome>;
}

function LegacyEditorToolbar({
  children,
  className,
  style,
  ...toolbarProps
}: EditorToolbarProps & { children: ReactNode; style?: CSSProperties }) {
  return (
    <ToolbarModeContext.Provider value="legacy">
      <EditorToolbarContext.Provider value={toolbarProps}>
        <div
          className={className}
          data-testid="pptx-editor-toolbar"
          style={{ ...rootStyle, ...style }}
        >
          {children}
        </div>
      </EditorToolbarContext.Provider>
    </ToolbarModeContext.Provider>
  );
}

/**
 * The toolbar region. With `mode="commands"` its parts read and run the
 * editor's commands, inside `PptxEditor`'s `toolbar` or under a
 * `PptxCommandProvider`; otherwise it provides the legacy props to
 * `EditorToolbar.Toolbar` and `useEditorToolbar()`.
 */
function EditorToolbarBase(
  props:
    | (EditorToolbarProps & { children: ReactNode; style?: CSSProperties })
    | EditorToolbarCommandProps
) {
  return props.mode === 'commands' ? (
    <CommandEditorToolbar {...props} />
  ) : (
    <LegacyEditorToolbar {...props} />
  );
}

const EditorToolbar = EditorToolbarBase as EditorToolbarComponent;
EditorToolbar.Toolbar = Toolbar;

export { EditorToolbar };
