import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import { commandMenuItem, menuSections } from '../../../../shared/react-toolbar/ContextMenu';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/MenuList';
import { columnLetter } from '../columnLetter';
import { commandLabelKey } from '../commands/descriptors';
import { storeAriaShortcut, storeShortcutLabel } from '../commands/shortcuts';
import type { XlsxCommandArgs, XlsxCommandId, XlsxCommandStore } from '../commands/types';

/** What a context menu opened on. */
export type XlsxContextMenuTarget = 'cell' | 'columnHeader' | 'rowHeader' | 'sheetTab';

/** A named group of a built-in menu, so later additions can join one by name. */
export interface XlsxContextMenuSection {
  id: string;
  entries: OverflowMenuEntry[];
}

/** The frozen panes a header menu can change, and the track it opened on. */
export interface XlsxContextFreezeTarget {
  /** Frozen track counts currently painted for the sheet. */
  rows: number;
  cols: number;
  /** The 0-based sheet column (columnHeader) or row (rowHeader) the menu opened on. */
  index: number;
  /** Apply a split: rows above and columns left of the given counts; 0/0 clears. */
  apply(rows: number, cols: number): void;
}

type Translate = (key: TranslationKey, vars?: Record<string, string | number>) => string;

/** The view entries freezing to the pressed header or clearing the split. */
function freezeEntries(
  columns: boolean,
  t: Translate,
  freeze: XlsxContextFreezeTarget
): OverflowMenuEntry[] {
  const entries: OverflowMenuEntry[] = [];
  if (columns) {
    const left = freeze.index;
    if (left > 0 && freeze.cols !== left) {
      entries.push(commandMenuItem({
        id: `freezeColumns:${left}`,
        label: left === 1
          ? t('commands.freezeColumn', { column: columnLetter(0) })
          : t('commands.freezeColumns', { first: columnLetter(0), last: columnLetter(left - 1) }),
        state: { enabled: true },
        run: () => freeze.apply(freeze.rows, left),
      }));
    }
  } else {
    const above = freeze.index;
    if (above > 0 && freeze.rows !== above) {
      entries.push(commandMenuItem({
        id: `freezeRows:${above}`,
        label: above === 1
          ? t('commands.freezeRow', { row: 1 })
          : t('commands.freezeRows', { first: 1, last: above }),
        state: { enabled: true },
        run: () => freeze.apply(above, freeze.cols),
      }));
    }
  }
  if (freeze.rows > 0 || freeze.cols > 0) {
    entries.push(commandMenuItem({
      id: 'unfreezePanes',
      label: t('commands.unfreezePanes'),
      state: { enabled: true },
      run: () => freeze.apply(0, 0),
    }));
  }
  return entries;
}

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
  t: Translate,
  freeze?: XlsxContextFreezeTarget
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
  if (target === 'columnHeader' || target === 'rowHeader') {
    const columns = target === 'columnHeader';
    const sections: XlsxContextMenuSection[] = [
      { id: 'cells', entries: columns ? [
        item(store, t, 'insertColumns', { position: 'left' }),
        item(store, t, 'insertColumns', { position: 'right' }),
      ] : [
        item(store, t, 'insertRows', { position: 'above' }),
        item(store, t, 'insertRows', { position: 'below' }),
      ] },
      { id: 'delete', entries: [
        columns ? item(store, t, 'deleteColumns', null) : item(store, t, 'deleteRows', null),
      ] },
    ];
    const view = freeze ? freezeEntries(columns, t, freeze) : [];
    if (view.length > 0) sections.push({ id: 'view', entries: view });
    return sections;
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
  t: Translate,
  freeze?: XlsxContextFreezeTarget
): OverflowMenuEntry[] {
  return menuSections(xlsxContextMenuSections(target, store, t, freeze).map((section) => section.entries));
}
