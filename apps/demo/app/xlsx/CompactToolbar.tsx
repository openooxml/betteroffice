"use client";

import {
  EditorToolbar,
  ToolbarButton,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  ToolbarGroup,
  XlsxPluginToolbar,
} from "@betteroffice/xlsx-react";

/** A task-focused toolbar composed only from public parts, in the host's order. */
export function CompactToolbar({ onShare }: { onShare(): void }) {
  return (
    <EditorToolbar mode="commands">
      <EditorToolbar.Toolbar>
        <ToolbarGroup label="Formatting">
          <ToolbarCommandSelect id="numberFormat" />
          <ToolbarCommandButton id="bold" />
        </ToolbarGroup>
        <ToolbarCommandButton id="undo" />
        <XlsxPluginToolbar />
        <ToolbarButton title="Copy link" onClick={onShare}>
          Copy link
        </ToolbarButton>
      </EditorToolbar.Toolbar>
      <EditorToolbar.FormulaBar />
    </EditorToolbar>
  );
}
