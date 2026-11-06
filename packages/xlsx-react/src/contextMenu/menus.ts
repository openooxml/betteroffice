import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import { commandMenuItem, menuSections } from '../../../../shared/react-toolbar/ContextMenu';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/MenuList';
import { commandLabelKey } from '../commands/descriptors';
import { storeAriaShortcut, storeShortcutLabel } from '../commands/shortcuts';
import type { XlsxCommandArgs, XlsxCommandId, XlsxCommandStore } from '../commands/types';

/** What a context menu opened on. */
export type XlsxContextMenuTarget = 'cell' | 'sheetTab';

/** A named group of a built-in menu, so later additions can join one by name. */
export interface XlsxContextMenuSection {
  id: string;
  entries: OverflowMenuEntry[];
}

type Translate = (key: TranslationKey) => string;

function item<K extends XlsxCommandId>(
  store: XlsxCommandStore,
  t: Translate,
  id: K,
  args: XlsxCommandArgs[K]
): OverflowMenuEntry {
  const bound = args === null ? undefined : args;
  return commandMenuItem({
    id: bound ? `${id}:${JSON.stringify(bound)}` : id,
    label: t(commandLabelKey(id, bound)),
    state: store.getState(id, args),
    shortcut: storeShortcutLabel(store, id, t, bound),
    ariaShortcut: storeAriaShortcut(store, id, bound),
    run: () => { void store.execute(id, args); },
  });
}

function submenu(id: string, label: string, entries: OverflowMenuEntry[]): OverflowMenuEntry {
  const disabled = entries.every((entry) => entry.kind === 'item' && entry.disabled);
  return { kind: 'submenu', id, label, entries, disabled };
}

/** The sections of a built-in menu, in Excel's order. */
export function xlsxContextMenuSections(
  target: XlsxContextMenuTarget,
  store: XlsxCommandStore,
  t: Translate
): XlsxContextMenuSection[] {
  if (target === 'sheetTab') {
    return [
      { id: 'sheet', entries: [
        item(store, t, 'insertSheet', null),
        item(store, t, 'deleteSheet', null),
        item(store, t, 'renameSheet', null),
      ] },
      { id: 'arrange', entries: [
        item(store, t, 'moveSheet', { direction: 'left' }),
        item(store, t, 'moveSheet', { direction: 'right' }),
      ] },
    ];
  }
  const numberFormat = store.getState('numberFormat');
  return [
    { id: 'clipboard', entries: [
      item(store, t, 'cut', null),
      item(store, t, 'copy', null),
      item(store, t, 'paste', null),
      item(store, t, 'pastePlainText', null),
    ] },
    { id: 'cells', entries: [
      submenu('insert', t('commands.insert'), [
        item(store, t, 'insertRows', { position: 'above' }),
        item(store, t, 'insertRows', { position: 'below' }),
        item(store, t, 'insertColumns', { position: 'left' }),
        item(store, t, 'insertColumns', { position: 'right' }),
      ]),
      submenu('delete', t('commands.delete'), [
        item(store, t, 'deleteRows', null),
        item(store, t, 'deleteColumns', null),
      ]),
      item(store, t, 'clearContents', null),
    ] },
    { id: 'format', entries: numberFormat.options ? [
      submenu('numberFormat', t('commands.numberFormat'), numberFormat.options.filter((option) => option.args.value !== 'custom').map((option) => commandMenuItem({
        id: `numberFormat:${option.args.value}`,
        label: option.label,
        state: option.state,
        run: () => { void store.execute('numberFormat', option.args); },
      })).map((entry) => ({ ...entry, ...(entry.kind === 'item' ? { radio: true,
        checked: numberFormat.value?.kind === entry.id.slice('numberFormat:'.length) } : {}) }))),
    ] : [] },
  ];
}

/** The entries of a built-in menu, sections joined by separators. */
export function xlsxContextMenuEntries(
  target: XlsxContextMenuTarget,
  store: XlsxCommandStore,
  t: Translate
): OverflowMenuEntry[] {
  return menuSections(xlsxContextMenuSections(target, store, t).map((section) => section.entries));
}
