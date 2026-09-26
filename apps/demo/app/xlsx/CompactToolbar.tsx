"use client";

import {
  EditorToolbar,
  ToolbarButton,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  ToolbarGroup,
} from "@betteroffice/xlsx-react";

export function CompactToolbar({ onShare }: { onShare(): void }) {
  return (
    <EditorToolbar mode="commands">
      <EditorToolbar.Toolbar>
        <ToolbarGroup label="Formatting">
          <ToolbarCommandSelect id="numberFormat" />
          <ToolbarCommandButton id="bold" />
        </ToolbarGroup>
        <ToolbarCommandButton id="undo" />
        <ToolbarButton title="Copy link" onClick={onShare}>
          Copy link
        </ToolbarButton>
      </EditorToolbar.Toolbar>
      <EditorToolbar.FormulaBar />
    </EditorToolbar>
  );
}
