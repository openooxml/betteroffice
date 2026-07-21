import type { CSSProperties, ReactNode } from 'react';
import { EditorToolbarContext } from './EditorToolbarContext';
import type { EditorToolbarProps } from './EditorToolbarContext';
import { Toolbar } from './Toolbar';

interface EditorToolbarComponent {
  (props: EditorToolbarProps & { children: ReactNode }): React.JSX.Element;
  Toolbar: typeof Toolbar;
}

function EditorToolbarBase({
  children,
  className,
  style,
  ...toolbarProps
}: EditorToolbarProps & { children: ReactNode; style?: CSSProperties }) {
  return (
    <EditorToolbarContext.Provider value={toolbarProps}>
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
    </EditorToolbarContext.Provider>
  );
}

const EditorToolbar = EditorToolbarBase as EditorToolbarComponent;
EditorToolbar.Toolbar = Toolbar;

export { EditorToolbar };
