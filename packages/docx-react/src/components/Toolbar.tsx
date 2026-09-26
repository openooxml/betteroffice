/**
 * The formatting rail. Without children it renders the default arrangement of
 * built-in controls; children replace that arrangement. Every control is bound
 * to the editor's command store, so the default and host-composed rails share
 * one authority for state, gating and shortcuts.
 */

import type { CSSProperties, ReactNode } from 'react';
import { useDocxCommandState } from '../commands/hooks';
import { useTranslation } from '../i18n';
import { useEditorChrome } from './EditorToolbarContext';
import { ToolbarCommand } from './toolbar/ToolbarCommand';
import { ToolbarRail } from './toolbar/ToolbarOverflow';
import { ToolbarGroup, ToolbarSeparator } from './toolbar/ToolbarPrimitives';

/** @experimental */
export interface ToolbarProps {
  /** The complete arrangement; omit to render the default controls. */
  children?: ReactNode;
  className?: string;
  style?: CSSProperties;
}

function DefaultToolbarItems() {
  const { t } = useTranslation();
  const chrome = useEditorChrome();
  const table = useDocxCommandState('tableAction');
  const image = useDocxCommandState('imageWrap');
  const inTable = table.value != null;
  const imageSelected = image.value != null;
  return (
    <>
      <ToolbarGroup label={t('formattingBar.groups.history')}>
        <ToolbarCommand id="undo" />
        <ToolbarCommand id="redo" />
      </ToolbarGroup>
      {(chrome?.showZoomControl ?? true) && (
        <ToolbarGroup label={t('formattingBar.groups.zoom')}>
          <ToolbarCommand id="zoom" />
        </ToolbarGroup>
      )}
      <ToolbarGroup label={t('formattingBar.groups.styles')}>
        <ToolbarCommand id="paragraphStyle" />
      </ToolbarGroup>
      <ToolbarGroup label={t('formattingBar.groups.font')}>
        <ToolbarCommand id="fontFamily" />
        <ToolbarCommand id="fontSize" />
      </ToolbarGroup>
      <ToolbarGroup label={t('formattingBar.groups.textFormatting')}>
        <ToolbarCommand id="bold" />
        <ToolbarCommand id="italic" />
        <ToolbarCommand id="underline" />
        <ToolbarCommand id="strikethrough" />
        <ToolbarCommand id="textColor" />
        <ToolbarCommand id="highlightColor" />
        <ToolbarCommand id="insertLink" />
      </ToolbarGroup>
      <ToolbarGroup label={t('formattingBar.groups.script')}>
        <ToolbarCommand id="superscript" />
        <ToolbarCommand id="subscript" />
      </ToolbarGroup>
      <ToolbarGroup label={t('formattingBar.groups.alignment')}>
        <ToolbarCommand id="alignment" />
      </ToolbarGroup>
      <ToolbarGroup label={t('formattingBar.groups.listFormatting')}>
        <ToolbarCommand id="bulletList" />
        <ToolbarCommand id="numberedList" />
        <ToolbarCommand id="outdent" />
        <ToolbarCommand id="indent" />
        <ToolbarCommand id="lineSpacing" />
      </ToolbarGroup>
      {imageSelected && (
        <ToolbarGroup label={t('formattingBar.groups.image')}>
          <ToolbarCommand id="imageWrap" />
          <ToolbarCommand id="imageTransform" />
          <ToolbarCommand id="imageProperties" />
        </ToolbarGroup>
      )}
      {inTable && (
        <ToolbarGroup label={t('formattingBar.groups.table')}>
          <ToolbarCommand id="tableAction" />
        </ToolbarGroup>
      )}
      <ToolbarCommand id="clearFormatting" />
      <ToolbarSeparator />
      <ToolbarCommand id="commentsSidebar" />
      <ToolbarSeparator />
      <ToolbarCommand id="editingMode" />
      {chrome?.toolbarExtra}
    </>
  );
}

/** The formatting rail, with overflow into an accessible "More" menu. */
export function Toolbar({ children, className, style }: ToolbarProps) {
  return (
    <ToolbarRail className={className} style={style} testId="formatting-bar">
      {children ?? <DefaultToolbarItems />}
    </ToolbarRail>
  );
}

export default Toolbar;
