/**
 * Keyboard-shortcut catalog — categorized list of every shortcut surfaced
 * in the KeyboardShortcutsDialog, plus the lookup helpers that filter and
 * label them.
 */

import type { TranslationKey } from '@betteroffice/docx-i18n';
import { DOCX_COMMAND_DESCRIPTORS, formatChord } from '../../../commands/descriptors';
import type {
  DocxCommandArgs,
  DocxCommandId,
  DocxCommandShortcut,
} from '../../../commands/types';
import type { KeyboardShortcut, ShortcutCategory } from '../KeyboardShortcutsDialog';

/**
 * Category label translation keys
 */
export const CATEGORY_LABEL_KEYS: Record<ShortcutCategory, TranslationKey> = {
  editing: 'dialogs.keyboardShortcuts.categories.editing',
  formatting: 'dialogs.keyboardShortcuts.categories.formatting',
  navigation: 'dialogs.keyboardShortcuts.categories.navigation',
  clipboard: 'dialogs.keyboardShortcuts.categories.clipboard',
  selection: 'dialogs.keyboardShortcuts.categories.selection',
  view: 'dialogs.keyboardShortcuts.categories.view',
  file: 'dialogs.keyboardShortcuts.categories.file',
  other: 'dialogs.keyboardShortcuts.categories.other',
};

/**
 * Category order for display
 */
export const CATEGORY_ORDER: ShortcutCategory[] = [
  'file',
  'editing',
  'clipboard',
  'formatting',
  'selection',
  'navigation',
  'view',
  'other',
];

interface CommandHelp<K extends DocxCommandId = DocxCommandId> {
  id: string;
  command: K;
  args?: DocxCommandArgs[K];
  name: string;
  nameKey: TranslationKey;
  description: string;
  descriptionKey: TranslationKey;
  category: ShortcutCategory;
  common?: boolean;
}

const help = (entry: CommandHelp): CommandHelp => entry;

/** Help text of editor commands; their keys come from the command descriptors. */
const COMMAND_HELP: readonly CommandHelp[] = [
  help({
    id: 'save',
    command: 'save',
    name: 'Save',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.save',
    description: 'Save document',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.saveDescription',
    category: 'file',
    common: true,
  }),
  help({
    id: 'open',
    command: 'open',
    name: 'Open',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.open',
    description: 'Open a document',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.openDescription',
    category: 'file',
  }),
  help({
    id: 'print',
    command: 'print',
    name: 'Print',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.print',
    description: 'Print document',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.printDescription',
    category: 'file',
  }),
  help({
    id: 'undo',
    command: 'undo',
    name: 'Undo',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.undo',
    description: 'Undo last action',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.undoDescription',
    category: 'editing',
    common: true,
  }),
  help({
    id: 'redo',
    command: 'redo',
    name: 'Redo',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.redo',
    description: 'Redo last action',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.redoDescription',
    category: 'editing',
    common: true,
  }),
  help({
    id: 'find',
    command: 'find',
    name: 'Find',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.find',
    description: 'Find text in document',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.findDescription',
    category: 'editing',
    common: true,
  }),
  help({
    id: 'replace',
    command: 'replace',
    name: 'Find & Replace',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.findReplace',
    description: 'Find and replace text',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.findReplaceDescription',
    category: 'editing',
  }),
  help({
    id: 'insert-link',
    command: 'insertLink',
    name: 'Insert Link',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.insertLink',
    description: 'Insert or edit a hyperlink',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.insertLinkDescription',
    category: 'editing',
  }),
  help({
    id: 'bold',
    command: 'bold',
    name: 'Bold',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.bold',
    description: 'Toggle bold formatting',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.boldDescription',
    category: 'formatting',
    common: true,
  }),
  help({
    id: 'italic',
    command: 'italic',
    name: 'Italic',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.italic',
    description: 'Toggle italic formatting',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.italicDescription',
    category: 'formatting',
    common: true,
  }),
  help({
    id: 'underline',
    command: 'underline',
    name: 'Underline',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.underline',
    description: 'Toggle underline formatting',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.underlineDescription',
    category: 'formatting',
    common: true,
  }),
  help({
    id: 'strikethrough',
    command: 'strikethrough',
    name: 'Strikethrough',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.strikethrough',
    description: 'Toggle strikethrough',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.strikethroughDescription',
    category: 'formatting',
  }),
  help({
    id: 'subscript',
    command: 'subscript',
    name: 'Subscript',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.subscript',
    description: 'Toggle subscript',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.subscriptDescription',
    category: 'formatting',
  }),
  help({
    id: 'superscript',
    command: 'superscript',
    name: 'Superscript',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.superscript',
    description: 'Toggle superscript',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.superscriptDescription',
    category: 'formatting',
  }),
  help({
    id: 'align-left',
    command: 'alignment',
    args: { value: 'left' },
    name: 'Align Left',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.alignLeft',
    description: 'Left align paragraph',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.alignLeftDescription',
    category: 'formatting',
  }),
  help({
    id: 'align-center',
    command: 'alignment',
    args: { value: 'center' },
    name: 'Align Center',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.alignCenter',
    description: 'Center align paragraph',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.alignCenterDescription',
    category: 'formatting',
  }),
  help({
    id: 'align-right',
    command: 'alignment',
    args: { value: 'right' },
    name: 'Align Right',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.alignRight',
    description: 'Right align paragraph',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.alignRightDescription',
    category: 'formatting',
  }),
  help({
    id: 'align-justify',
    command: 'alignment',
    args: { value: 'both' },
    name: 'Justify',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.justify',
    description: 'Justify paragraph',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.justifyDescription',
    category: 'formatting',
  }),
  help({
    id: 'indent',
    command: 'indent',
    name: 'Increase Indent',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.increaseIndent',
    description: 'Increase paragraph indent',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.increaseIndentDescription',
    category: 'formatting',
  }),
  help({
    id: 'outdent',
    command: 'outdent',
    name: 'Decrease Indent',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.decreaseIndent',
    description: 'Decrease paragraph indent',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.decreaseIndentDescription',
    category: 'formatting',
  }),
];

/** Shortcuts of editor commands, as bound by the command descriptors. */
export const COMMAND_SHORTCUTS: KeyboardShortcut[] = COMMAND_HELP.flatMap(
  ({ command, args, ...entry }) => {
    const key = JSON.stringify(args ?? null);
    const chords = (DOCX_COMMAND_DESCRIPTORS[command].shortcuts as readonly DocxCommandShortcut[])
      .filter((shortcut) => JSON.stringify(shortcut.args) === key)
      .map((shortcut) => formatChord(shortcut.chord, false));
    if (chords.length === 0) return [];
    return [{ ...entry, keys: chords[0], ...(chords[1] ? { altKeys: chords[1] } : {}) }];
  }
);

/** Keys the text input and the browser handle, outside the editor commands. */
const INPUT_SHORTCUTS: KeyboardShortcut[] = [
  {
    id: 'delete',
    name: 'Delete',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.delete',
    description: 'Delete selected text',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.deleteDescription',
    keys: 'Del',
    altKeys: 'Backspace',
    category: 'editing',
  },

  // Clipboard
  {
    id: 'cut',
    name: 'Cut',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.cut',
    description: 'Cut selected text',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.cutDescription',
    keys: 'Ctrl+X',
    category: 'clipboard',
    common: true,
  },
  {
    id: 'copy',
    name: 'Copy',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.copy',
    description: 'Copy selected text',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.copyDescription',
    keys: 'Ctrl+C',
    category: 'clipboard',
    common: true,
  },
  {
    id: 'paste',
    name: 'Paste',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.paste',
    description: 'Paste from clipboard',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.pasteDescription',
    keys: 'Ctrl+V',
    category: 'clipboard',
    common: true,
  },
  {
    id: 'paste-plain',
    name: 'Paste as Plain Text',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.pastePlainText',
    description: 'Paste without formatting',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.pastePlainTextDescription',
    keys: 'Ctrl+Shift+V',
    category: 'clipboard',
  },

  // Selection
  {
    id: 'select-all',
    name: 'Select All',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.selectAll',
    description: 'Select all content',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.selectAllDescription',
    keys: 'Ctrl+A',
    category: 'selection',
    common: true,
  },
  {
    id: 'select-word',
    name: 'Select Word',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.selectWord',
    description: 'Select current word',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.selectWordDescription',
    keys: 'Double-click',
    category: 'selection',
  },
  {
    id: 'select-paragraph',
    name: 'Select Paragraph',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.selectParagraph',
    description: 'Select current paragraph',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.selectParagraphDescription',
    keys: 'Triple-click',
    category: 'selection',
  },
  {
    id: 'extend-selection-word',
    name: 'Extend Selection by Word',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.extendSelectionByWord',
    description: 'Extend selection to next/previous word',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.extendSelectionByWordDescription',
    keys: 'Ctrl+Shift+Arrow',
    category: 'selection',
  },
  {
    id: 'extend-selection-line',
    name: 'Extend Selection to Line Edge',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.extendSelectionToLineEdge',
    description: 'Extend selection to line start/end',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.extendSelectionToLineEdgeDescription',
    keys: 'Shift+Home/End',
    category: 'selection',
  },

  // Navigation
  {
    id: 'move-word',
    name: 'Move by Word',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.moveByWord',
    description: 'Move cursor to next/previous word',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.moveByWordDescription',
    keys: 'Ctrl+Arrow',
    category: 'navigation',
  },
  {
    id: 'move-line-start',
    name: 'Move to Line Start',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.moveToLineStart',
    description: 'Move cursor to start of line',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.moveToLineStartDescription',
    keys: 'Home',
    category: 'navigation',
  },
  {
    id: 'move-line-end',
    name: 'Move to Line End',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.moveToLineEnd',
    description: 'Move cursor to end of line',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.moveToLineEndDescription',
    keys: 'End',
    category: 'navigation',
  },
  {
    id: 'move-doc-start',
    name: 'Move to Document Start',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.moveToDocumentStart',
    description: 'Move cursor to start of document',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.moveToDocumentStartDescription',
    keys: 'Ctrl+Home',
    category: 'navigation',
  },
  {
    id: 'move-doc-end',
    name: 'Move to Document End',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.moveToDocumentEnd',
    description: 'Move cursor to end of document',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.moveToDocumentEndDescription',
    keys: 'Ctrl+End',
    category: 'navigation',
  },
  {
    id: 'page-up',
    name: 'Page Up',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.pageUp',
    description: 'Scroll up one page',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.pageUpDescription',
    keys: 'Page Up',
    category: 'navigation',
  },
  {
    id: 'page-down',
    name: 'Page Down',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.pageDown',
    description: 'Scroll down one page',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.pageDownDescription',
    keys: 'Page Down',
    category: 'navigation',
  },

  // View
  {
    id: 'shortcuts',
    name: 'Keyboard Shortcuts',
    nameKey: 'dialogs.keyboardShortcuts.shortcuts.keyboardShortcuts',
    description: 'Show this help dialog',
    descriptionKey: 'dialogs.keyboardShortcuts.shortcuts.keyboardShortcutsDescription',
    keys: 'Ctrl+/',
    altKeys: 'F1',
    category: 'view',
  },
];

/**
 * Default keyboard shortcuts (with translation keys for name/description)
 */
export const DEFAULT_SHORTCUTS: KeyboardShortcut[] = [...COMMAND_SHORTCUTS, ...INPUT_SHORTCUTS];

/**
 * Get all default shortcuts
 */
export function getDefaultShortcuts(): KeyboardShortcut[] {
  return [...DEFAULT_SHORTCUTS];
}

/**
 * Get shortcuts by category
 */
export function getShortcutsByCategory(category: ShortcutCategory): KeyboardShortcut[] {
  return DEFAULT_SHORTCUTS.filter((s) => s.category === category);
}

/**
 * Get common/frequently used shortcuts
 */
export function getCommonShortcuts(): KeyboardShortcut[] {
  return DEFAULT_SHORTCUTS.filter((s) => s.common);
}

/**
 * Get category label translation key
 */
export function getCategoryLabel(category: ShortcutCategory): string {
  return CATEGORY_LABEL_KEYS[category];
}

/**
 * Get all categories
 */
export function getAllCategories(): ShortcutCategory[] {
  return [...CATEGORY_ORDER];
}
