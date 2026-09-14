import type { TFunction } from '@betteroffice/vsdx-i18n';
import { CommandMenu } from './CommandMenu';
import type { CommandMenuEntry } from './CommandMenu';
import type { RibbonCommandId } from './commands';

export interface ShapeContextMenuProps {
  t: TFunction;
  position: { top: number; left: number };
  onClose: () => void;
  onCloseAndFocus: () => void;
}

/** Shape operations offered on right-click, in Visio's relative order. */
export const SHAPE_CONTEXT_ENTRIES: ReadonlyArray<CommandMenuEntry> = [
  { id: 'delete', icon: 'delete' },
  {
    id: 'bringToFront', icon: 'front', children: [
      { id: 'bringToFront', icon: 'front' },
      { id: 'bringForward', icon: 'forward' },
    ],
  },
  {
    id: 'sendToBack', icon: 'back', children: [
      { id: 'sendBackward', icon: 'backward' },
      { id: 'sendToBack', icon: 'back' },
    ],
  },
];

const SHAPE_CONTEXT_DIVIDERS: ReadonlySet<RibbonCommandId> = new Set(['delete']);

/** Right-click menu for the selected shape, sharing the ribbon menus' keyboard behaviour. */
export function ShapeContextMenu({ t, position, onClose, onCloseAndFocus }: ShapeContextMenuProps) {
  return <CommandMenu menuLabel={t('contextMenu.label')} entries={SHAPE_CONTEXT_ENTRIES} position={position} dividerAfter={SHAPE_CONTEXT_DIVIDERS} label={(id) => t(`ribbon.commands.${id}`)} onClose={onClose} onCloseAndFocus={onCloseAndFocus} />;
}
