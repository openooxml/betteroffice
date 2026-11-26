import { useDocxCommand } from '../commands/hooks';
import type { DocxCommandArgs, DocxCommandId, DocxCommandState } from '../commands/types';
import { useEditorChrome } from './EditorToolbarContext';
import type { MenuEntry, MenuItem } from './ui/MenuDropdown';

export interface BoundMenuItem {
  entry: MenuItem;
  hidden: boolean;
}

/** Whether a menu leaves the command out rather than showing it disabled. */
export function hiddenInMenu(state: DocxCommandState): boolean {
  return (
    !state.enabled &&
    (state.disabledReason.code === 'host-disabled' ||
      state.disabledReason.code === 'unsupported-command')
  );
}

export function useCommandMenuItem<K extends DocxCommandId>(
  id: K,
  icon: string | undefined,
  args?: DocxCommandArgs[K]
): BoundMenuItem {
  const command = useDocxCommand(id, args);
  const prepareSave = useEditorChrome()?.prepareSave;
  const { state } = command;
  return {
    entry: {
      icon,
      label: command.label,
      shortcut: command.shortcut ?? undefined,
      disabled: !state.enabled,
      description: state.enabled ? undefined : state.disabledReason.message,
      onClick: () => void command.execute(),
      onPointerEnter: id === 'save' ? prepareSave : undefined,
      onFocus: id === 'save' ? prepareSave : undefined,
    },
    hidden: hiddenInMenu(state as DocxCommandState),
  };
}

export function visible(...items: BoundMenuItem[]): MenuEntry[] {
  return items.filter((item) => !item.hidden).map((item) => item.entry);
}
