"use client";

import {
  EditorToolbar,
  ToolbarButton,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  ToolbarGroup,
} from "@betteroffice/docx-react";

export function CompactToolbar({ onShare }: { onShare(): void }) {
  return (
    <EditorToolbar>
      <EditorToolbar.Toolbar>
        <ToolbarGroup label="Formatting">
          <ToolbarCommandSelect id="paragraphStyle" />
          <ToolbarCommandButton id="bold" />
        </ToolbarGroup>
        <ToolbarCommandButton id="undo" />
        <EditorToolbar.Review />
        <ToolbarButton title="Copy link" onClick={onShare}>
          Copy link
        </ToolbarButton>
      </EditorToolbar.Toolbar>
    </EditorToolbar>
  );
}
