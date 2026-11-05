/**
 * Text Context Menu Component
 *
 * Right-click context menu for text editing operations.
 * Shows Cut, Copy, Paste, and other text editing options.
 */

import React, { useCallback, useEffect, useState, useMemo } from 'react';
import type { ImageLayoutTarget } from '@betteroffice/docx/docx';
import { useTranslation } from '../i18n';
import type { TranslationKey } from '@betteroffice/docx-i18n';
import { ContextMenu } from '../../../../shared/react-toolbar/ContextMenu';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/MenuList';

// ============================================================================
// TYPES
// ============================================================================

/**
 * Context menu action types
 */
export type TextContextAction =
  | 'cut'
  | 'copy'
  | 'paste'
  | 'pasteAsPlainText'
  | 'selectAll'
  | 'delete'
  | 'separator'
  | 'addRowAbove'
  | 'addRowBelow'
  | 'deleteRow'
  | 'addColumnLeft'
  | 'addColumnRight'
  | 'deleteColumn'
  | 'mergeCells'
  | 'splitCell'
  | 'selectTable'
  | 'deleteTable'
  | 'addComment'
  | 'resizeImage'
  | 'moveImage'
  | 'resizeColumn'
  | `wrap:${ImageLayoutTarget}`;

/**
 * Menu item configuration
 */
export interface TextContextMenuItem {
  /** Action type */
  action: TextContextAction;
  /** Display label */
  label: string;
  /** Keyboard shortcut hint */
  shortcut?: string;
  /** Whether the item is disabled */
  disabled?: boolean;
  /** Whether to show divider after this item */
  dividerAfter?: boolean;
}

/**
 * Context menu props
 */
export interface TextContextMenuProps {
  /** Whether the menu is visible */
  isOpen: boolean;
  /** Menu position */
  position: { x: number; y: number };
  /** Whether there's a selection (enables copy/cut) */
  hasSelection: boolean;
  /** Whether the editor is editable (enables paste/cut/delete) */
  isEditable: boolean;
  /** Whether clipboard has content (enables paste) */
  hasClipboardContent?: boolean;
  /** Callback when an action is selected */
  onAction: (action: TextContextAction) => void;
  /** Callback when menu is closed */
  onClose: () => void;
  /** Custom menu items (overrides default) */
  items?: TextContextMenuItem[];
  /** Additional className */
  className?: string;
  /** Moves focus back to the document when the menu closes from the keyboard or an item */
  returnFocus?: () => void;
}

/**
 * Hook options for text context menu
 */
export interface UseTextContextMenuOptions {
  /** Whether the context menu is enabled */
  enabled?: boolean;
  /** Whether the editor is editable */
  isEditable?: boolean;
  /** Container element ref */
  containerRef?: React.RefObject<HTMLElement>;
  /** Callback when an action is triggered */
  onAction?: (action: TextContextAction) => void;
}

/**
 * Hook return value
 */
export interface UseTextContextMenuReturn {
  /** Whether the menu is open */
  isOpen: boolean;
  /** Menu position */
  position: { x: number; y: number };
  /** Whether there's a text selection */
  hasSelection: boolean;
  /** Open the context menu */
  openMenu: (event: React.MouseEvent | MouseEvent) => void;
  /** Close the context menu */
  closeMenu: () => void;
  /** Handle action selection */
  handleAction: (action: TextContextAction) => void;
  /** Context menu event handler for onContextMenu prop */
  onContextMenu: (event: React.MouseEvent) => void;
}

// ============================================================================
// CONSTANTS
// ============================================================================

/**
 * Default menu item definitions (use translation keys for label/shortcut)
 */
interface DefaultMenuItemDef {
  action: TextContextAction;
  labelKey: TranslationKey;
  shortcutKey?: TranslationKey;
  dividerAfter?: boolean;
}

const DEFAULT_MENU_ITEM_DEFS: DefaultMenuItemDef[] = [
  { action: 'cut', labelKey: 'contextMenu.cut', shortcutKey: 'contextMenu.cutShortcut' },
  { action: 'copy', labelKey: 'contextMenu.copy', shortcutKey: 'contextMenu.copyShortcut' },
  { action: 'paste', labelKey: 'contextMenu.paste', shortcutKey: 'contextMenu.pasteShortcut' },
  {
    action: 'pasteAsPlainText',
    labelKey: 'contextMenu.pastePlainText',
    shortcutKey: 'contextMenu.pastePlainTextShortcut',
    dividerAfter: true,
  },
  {
    action: 'delete',
    labelKey: 'contextMenu.delete',
    shortcutKey: 'contextMenu.deleteShortcut',
    dividerAfter: true,
  },
  {
    action: 'selectAll',
    labelKey: 'contextMenu.selectAll',
    shortcutKey: 'contextMenu.selectAllShortcut',
  },
];

// ============================================================================
// ICONS
// ============================================================================

const CutIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <circle cx="4" cy="12" r="2" stroke="currentColor" strokeWidth="1.5" />
    <circle cx="12" cy="12" r="2" stroke="currentColor" strokeWidth="1.5" />
    <path
      d="M5.5 10.5L10.5 3M10.5 10.5L5.5 3"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
    />
  </svg>
);

const CopyIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="5" y="5" width="8" height="9" rx="1" stroke="currentColor" strokeWidth="1.5" />
    <path
      d="M11 5V3a1 1 0 00-1-1H4a1 1 0 00-1 1v8a1 1 0 001 1h2"
      stroke="currentColor"
      strokeWidth="1.5"
    />
  </svg>
);

const PasteIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="3" y="3" width="10" height="11" rx="1" stroke="currentColor" strokeWidth="1.5" />
    <path d="M6 3V2a1 1 0 011-1h2a1 1 0 011 1v1" stroke="currentColor" strokeWidth="1.5" />
    <path d="M6 8h4M6 11h4" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </svg>
);

const DeleteIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </svg>
);

const SelectAllIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect
      x="2"
      y="2"
      width="12"
      height="12"
      rx="1"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeDasharray="2 2"
    />
    <rect x="4" y="4" width="8" height="8" fill="currentColor" opacity="0.3" />
  </svg>
);

const AddRowAboveIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="2" y="6" width="12" height="4" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <rect x="2" y="10" width="12" height="4" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <path d="M8 1v3M6.5 2.5h3" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
  </svg>
);

const AddRowBelowIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="2" y="2" width="12" height="4" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <rect x="2" y="6" width="12" height="4" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <path d="M8 12v3M6.5 13.5h3" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
  </svg>
);

const DeleteRowIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="2" y="2" width="12" height="4" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <rect
      x="2"
      y="6"
      width="12"
      height="4"
      rx="0.5"
      stroke="currentColor"
      strokeWidth="1.2"
      opacity="0.3"
    />
    <rect x="2" y="10" width="12" height="4" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <path d="M5 8h6" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </svg>
);

const AddColumnLeftIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="6" y="2" width="4" height="12" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <rect x="10" y="2" width="4" height="12" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <path d="M3 8H0.5M1.75 6.5v3" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
  </svg>
);

const AddColumnRightIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="2" y="2" width="4" height="12" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <rect x="6" y="2" width="4" height="12" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <path d="M13 8h2.5M14.25 6.5v3" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
  </svg>
);

const DeleteColumnIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="2" y="2" width="4" height="12" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <rect
      x="6"
      y="2"
      width="4"
      height="12"
      rx="0.5"
      stroke="currentColor"
      strokeWidth="1.2"
      opacity="0.3"
    />
    <rect x="10" y="2" width="4" height="12" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <path d="M8 5v6" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </svg>
);

const MergeCellsIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="2" y="4" width="4" height="8" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <rect x="10" y="4" width="4" height="8" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <path d="M7 8h2" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
    <path d="M8 7l1 1-1 1" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
  </svg>
);

const SplitCellIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="2" y="4" width="12" height="8" rx="0.5" stroke="currentColor" strokeWidth="1.2" />
    <path d="M8 4v8" stroke="currentColor" strokeWidth="1.2" strokeDasharray="2 1" />
    <path d="M6.5 8h-1M9.5 8h1" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </svg>
);

const SelectTableIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect
      x="2"
      y="2"
      width="12"
      height="12"
      rx="1"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeDasharray="2 1.5"
    />
    <path d="M8 2v12M2 8h12" stroke="currentColor" strokeWidth="1.2" />
  </svg>
);

const DeleteTableIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="2" y="2" width="12" height="12" rx="1" stroke="currentColor" strokeWidth="1.3" />
    <path d="M5 2v12M2 8h11" stroke="currentColor" strokeWidth="1.1" opacity="0.6" />
    <path
      d="M5.5 5.5l5 5M10.5 5.5l-5 5"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
    />
  </svg>
);

const CommentIcon = () => (
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path
      d="M3 3h10a1 1 0 011 1v6a1 1 0 01-1 1H6l-3 2.5V11H3a1 1 0 01-1-1V4a1 1 0 011-1z"
      stroke="currentColor"
      strokeWidth="1.3"
    />
    <path d="M5 6h6M5 8.5h4" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
  </svg>
);

/**
 * Get icon for action
 */
function getActionIcon(action: TextContextAction): React.ReactNode {
  switch (action) {
    case 'cut':
      return <CutIcon />;
    case 'copy':
      return <CopyIcon />;
    case 'paste':
    case 'pasteAsPlainText':
      return <PasteIcon />;
    case 'delete':
      return <DeleteIcon />;
    case 'selectAll':
      return <SelectAllIcon />;
    case 'addRowAbove':
      return <AddRowAboveIcon />;
    case 'addRowBelow':
      return <AddRowBelowIcon />;
    case 'deleteRow':
      return <DeleteRowIcon />;
    case 'addColumnLeft':
      return <AddColumnLeftIcon />;
    case 'addColumnRight':
      return <AddColumnRightIcon />;
    case 'deleteColumn':
      return <DeleteColumnIcon />;
    case 'mergeCells':
      return <MergeCellsIcon />;
    case 'splitCell':
      return <SplitCellIcon />;
    case 'selectTable':
      return <SelectTableIcon />;
    case 'deleteTable':
      return <DeleteTableIcon />;
    case 'addComment':
      return <CommentIcon />;
    default:
      return null;
  }
}

// ============================================================================
// TEXT CONTEXT MENU COMPONENT
// ============================================================================

const MENU_COLORS = {
  surface: 'var(--doc-surface)',
  text: 'var(--doc-text)',
  mutedText: 'var(--doc-text-subtle)',
  border: 'var(--doc-border-light)',
  hover: 'var(--doc-primary-light)',
  shadow: 'var(--doc-shadow)',
  focus: 'var(--doc-primary)',
};

export const TextContextMenu: React.FC<TextContextMenuProps> = ({
  isOpen,
  position,
  hasSelection,
  isEditable,
  hasClipboardContent = true,
  onAction,
  onClose,
  items,
  className = '',
  returnFocus,
}) => {
  const { t } = useTranslation();

  const translatedDefaults: TextContextMenuItem[] = useMemo(
    () =>
      DEFAULT_MENU_ITEM_DEFS.map((def) => ({
        action: def.action,
        label: t(def.labelKey),
        shortcut: def.shortcutKey ? t(def.shortcutKey) : undefined,
        dividerAfter: def.dividerAfter,
      })),
    [t]
  );

  const entries: OverflowMenuEntry[] = [];
  (items || translatedDefaults).forEach((item, index) => {
    if (item.action === 'separator') {
      entries.push({ kind: 'separator', id: `separator-${index}` });
      return;
    }
    const disabled =
      item.disabled ??
      (item.action === 'cut' || item.action === 'copy' || item.action === 'delete'
        ? !hasSelection
        : item.action === 'paste' || item.action === 'pasteAsPlainText'
          ? !isEditable || !hasClipboardContent
          : false);
    entries.push({
      kind: 'item',
      id: `${item.action}-${index}`,
      label: item.label,
      icon: getActionIcon(item.action),
      shortcut: item.shortcut,
      disabled,
      onSelect: () => onAction(item.action),
    });
    if (item.dividerAfter) entries.push({ kind: 'separator', id: `divider-${index}` });
  });

  if (!isOpen) return null;

  return (
    <ContextMenu
      label={t('contextMenu.textMenuAriaLabel')}
      entries={entries}
      point={position}
      colors={MENU_COLORS}
      className={`docx-text-context-menu ${className}`.trim()}
      returnFocus={returnFocus}
      onClose={() => onClose()}
    />
  );
};

// ============================================================================
// USE TEXT CONTEXT MENU HOOK
// ============================================================================

/**
 * Hook to manage text context menu state
 */
export function useTextContextMenu(
  options: UseTextContextMenuOptions = {}
): UseTextContextMenuReturn {
  const {
    enabled = true,
    isEditable: _isEditable = true,
    containerRef: _containerRef,
    onAction,
  } = options;

  const [isOpen, setIsOpen] = useState(false);
  const [position, setPosition] = useState({ x: 0, y: 0 });
  const [hasSelection, setHasSelection] = useState(false);

  /**
   * Check if there's a text selection
   */
  const checkSelection = useCallback(() => {
    const selection = window.getSelection();
    const hasText = selection && !selection.isCollapsed && selection.toString().length > 0;
    setHasSelection(!!hasText);
    return !!hasText;
  }, []);

  /**
   * Open the context menu
   */
  const openMenu = useCallback(
    (event: React.MouseEvent | MouseEvent) => {
      if (!enabled) return;

      event.preventDefault();
      event.stopPropagation();

      // Update selection state
      checkSelection();

      setPosition({ x: event.clientX, y: event.clientY });
      setIsOpen(true);
    },
    [enabled, checkSelection]
  );

  /**
   * Close the context menu
   */
  const closeMenu = useCallback(() => {
    setIsOpen(false);
  }, []);

  /**
   * Handle action selection
   */
  const handleAction = useCallback(
    (action: TextContextAction) => {
      closeMenu();

      // Execute the action
      switch (action) {
        case 'cut':
          document.execCommand('cut');
          break;
        case 'copy':
          document.execCommand('copy');
          break;
        case 'paste':
          document.execCommand('paste');
          break;
        case 'pasteAsPlainText':
          // Trigger paste event with shift key simulation
          // Note: This may not work in all browsers due to security restrictions
          navigator.clipboard
            .readText?.()
            .then((text) => {
              document.execCommand('insertText', false, text);
            })
            .catch(() => {
              // Fallback - just try regular paste
              document.execCommand('paste');
            });
          break;
        case 'delete':
          document.execCommand('delete');
          break;
        case 'selectAll':
          document.execCommand('selectAll');
          break;
      }

      onAction?.(action);
    },
    [closeMenu, onAction]
  );

  /**
   * Context menu event handler
   */
  const onContextMenu = useCallback(
    (event: React.MouseEvent) => {
      openMenu(event);
    },
    [openMenu]
  );

  // Close menu when clicking elsewhere or pressing Escape
  useEffect(() => {
    if (!isOpen) return;

    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        closeMenu();
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, closeMenu]);

  return {
    isOpen,
    position,
    hasSelection,
    openMenu,
    closeMenu,
    handleAction,
    onContextMenu,
  };
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Get action shortcut
 */
export function getTextActionShortcut(action: TextContextAction): string {
  const shortcuts: Partial<Record<TextContextAction, string>> = {
    cut: 'Ctrl+X',
    copy: 'Ctrl+C',
    paste: 'Ctrl+V',
    pasteAsPlainText: 'Ctrl+Shift+V',
    selectAll: 'Ctrl+A',
    delete: 'Del',
    separator: '',
    addRowAbove: '',
    addRowBelow: '',
    deleteRow: '',
    addColumnLeft: '',
    addColumnRight: '',
    deleteColumn: '',
    mergeCells: '',
    splitCell: '',
    selectTable: '',
    deleteTable: '',
    addComment: '',
    resizeImage: '',
    moveImage: '',
    resizeColumn: '',
  };
  return shortcuts[action] ?? '';
}

/**
 * Get default menu item definitions (untranslated, use translation keys)
 */
export function getDefaultTextContextMenuItems(): TextContextMenuItem[] {
  return DEFAULT_MENU_ITEM_DEFS.map((def) => ({
    action: def.action,
    label: def.labelKey,
    shortcut: def.shortcutKey,
    dividerAfter: def.dividerAfter,
  }));
}

/**
 * Check if action is available
 */
export function isTextActionAvailable(
  action: TextContextAction,
  hasSelection: boolean,
  isEditable: boolean
): boolean {
  switch (action) {
    case 'cut':
    case 'copy':
    case 'delete':
      return hasSelection;
    case 'paste':
    case 'pasteAsPlainText':
      return isEditable;
    case 'addComment':
      return hasSelection;
    case 'selectAll':
      return true;
    default:
      return true;
  }
}

// ============================================================================
// EXPORTS
// ============================================================================

export default TextContextMenu;
