"use client";

import {
  EditorToolbar,
  ToolbarButton,
  ToolbarCommandButton,
  ToolbarCommandSelect,
  ToolbarGroup,
} from "@betteroffice/pptx-react";

export function CompactToolbar({ onShare }: { onShare(): void }) {
  return (
    <EditorToolbar mode="commands">
      <EditorToolbar.Toolbar>
        <ToolbarGroup label="Text">
          <ToolbarCommandSelect id="fontFamily" />
          <ToolbarCommandButton id="bold" />
        </ToolbarGroup>
        <ToolbarCommandButton id="undo" />
        <ToolbarCommandButton id="slideshow" />
        <ToolbarButton title="Copy link" onClick={onShare}>
          Copy link
        </ToolbarButton>
      </EditorToolbar.Toolbar>
    </EditorToolbar>
  );
}
