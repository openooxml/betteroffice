import type { TranslationKey } from '@betteroffice/pptx-i18n';
import { commandMenuItem, menuSections } from '../../../../shared/react-toolbar/ContextMenu';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/MenuList';
import { commandShortcut } from '../commands/descriptors';
import type { PptxCommandArgs, PptxCommandId, PptxCommandStore } from '../commands/types';

/** @internal */
export type PptxContextMenuTarget = 'text' | 'shape' | 'table' | 'slide' | 'empty';

type Translate = (key: TranslationKey) => string;

function item<K extends PptxCommandId>(store: PptxCommandStore, t: Translate, id: K,
  args: PptxCommandArgs[K], labelKey = store.getDescriptor(id).labelKey): OverflowMenuEntry {
  return commandMenuItem({
    id: `${id}:${JSON.stringify(args)}`,
    label: t(labelKey),
    state: store.getState(id, args),
    shortcut: commandShortcut(id, args),
    run: () => { void store.execute(id, args); },
  });
}

function submenu(id: string, label: string, entries: OverflowMenuEntry[]): OverflowMenuEntry {
  return { kind: 'submenu', id, label, entries,
    disabled: entries.every((entry) => entry.kind === 'item' && entry.disabled) };
}

/** @internal */
export function pptxContextMenuSections(target: PptxContextMenuTarget, store: PptxCommandStore,
  t: Translate): { id: string; entries: OverflowMenuEntry[] }[] {
  if (target === 'empty') return [];
  if (target === 'slide') return [{ id: 'slide', entries: [
    item(store, t, 'insertSlide', {}),
    item(store, t, 'duplicateSlide', null),
    item(store, t, 'deleteSlide', null),
  ] }];
  if (target === 'shape') return [
    { id: 'object', entries: [item(store, t, 'deleteObject', null)] },
    { id: 'arrange', entries: [
      submenu('front', t('toolbar.bringToFront'), [
        item(store, t, 'zOrder', { value: 'front' }, 'toolbar.bringToFront'),
        item(store, t, 'zOrder', { value: 'forward' }, 'toolbar.bringForward'),
      ]),
      submenu('back', t('toolbar.sendToBack'), [
        item(store, t, 'zOrder', { value: 'back' }, 'toolbar.sendToBack'),
        item(store, t, 'zOrder', { value: 'backward' }, 'toolbar.sendBackward'),
      ]),
    ] },
  ];
  const sections = [{ id: 'clipboard', entries: [
    item(store, t, 'cut', null), item(store, t, 'copy', null), item(store, t, 'paste', null),
  ] }];
  if (target === 'table') sections.push({ id: 'table', entries: [
    submenu('insert', t('commands.insert'), [
      item(store, t, 'tableInsert', { kind: 'row', position: 'before' }, 'commands.insertRowsAbove'),
      item(store, t, 'tableInsert', { kind: 'row', position: 'after' }, 'commands.insertRowsBelow'),
      item(store, t, 'tableInsert', { kind: 'column', position: 'before' }, 'commands.insertColumnsLeft'),
      item(store, t, 'tableInsert', { kind: 'column', position: 'after' }, 'commands.insertColumnsRight'),
    ]),
    submenu('delete', t('commands.delete'), [
      item(store, t, 'tableDelete', { kind: 'row' }, 'commands.deleteRows'),
      item(store, t, 'tableDelete', { kind: 'column' }, 'commands.deleteColumns'),
      item(store, t, 'tableDelete', { kind: 'table' }, 'commands.deleteTable'),
    ]),
  ] });
  return sections;
}

/** @internal */
export function pptxContextMenuEntries(target: PptxContextMenuTarget, store: PptxCommandStore,
  t: Translate): OverflowMenuEntry[] {
  return menuSections(pptxContextMenuSections(target, store, t).map((section) => section.entries));
}
