import type { TFunction } from '@betteroffice/vsdx-i18n';
import { CommandMenu } from './CommandMenu';
import type { CommandMenuEntry } from './CommandMenu';
import type { RibbonCommandId } from './commands';

export interface CanvasContextMenuProps {
  t: TFunction;
  position: { top: number; left: number };
  onClose: () => void;
  onCloseAndFocus: () => void;
}

/** Canvas operations offered on empty-canvas right-click, reusing ribbon commands. */
export const CANVAS_CONTEXT_ENTRIES: ReadonlyArray<CommandMenuEntry> = [
  { id: 'undo', icon: 'undo' },
  { id: 'redo', icon: 'redo' },
  { id: 'addShape', icon: 'add' },
];

const CANVAS_CONTEXT_DIVIDERS: ReadonlySet<RibbonCommandId> = new Set(['redo']);

/** Right-click menu for empty canvas, sharing the ribbon menus' keyboard behaviour. */
export function CanvasContextMenu({ t, position, onClose, onCloseAndFocus }: CanvasContextMenuProps) {
  return <CommandMenu menuLabel={t('contextMenu.canvasLabel')} entries={CANVAS_CONTEXT_ENTRIES} position={position} dividerAfter={CANVAS_CONTEXT_DIVIDERS} label={(id) => t(`ribbon.commands.${id}`)} onClose={onClose} onCloseAndFocus={onCloseAndFocus} />;
}
