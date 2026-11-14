import { useCallback, useMemo, useState } from 'react';
import type { ImageLayoutTarget } from '@betteroffice/docx/docx';
import type { WrapType } from '@betteroffice/docx/docx/wrapTypes';
import { en as defaultLocale } from '@betteroffice/docx-i18n';
import { useTranslation } from '../../../i18n';
import type { Translations } from '@betteroffice/docx-i18n';
import { useImageContextMenu } from '../../ImageContextMenu';
import { type TextContextAction, type TextContextMenuItem } from '../../TextContextMenu';
import {
  effectiveZoom,
  resolveDisplayPageClientRect,
  type DisplayListQueries,
} from '@betteroffice/docx/layout/render';
import { isWithinPageArea } from '../internals/pageAreaRouting';
import { formatKeys } from '../../dialogs/KeyboardShortcutsDialog/ShortcutItem';
import { storeAriaShortcut, storeShortcutLabel } from '../../../commands/shortcuts';
import { useDocxShortcutRevision } from '../../../commands/hooks';
import type { DocxCommandId, DocxCommandStore } from '../../../commands/types';
import type { PagedEditorRef } from '../PagedEditor';
import { currentYrsTableTarget } from '../yrsCommands';
import { readClipboardFragment } from '../internals/clipboardHtml';
import type { DocxTableAction } from '../../../commands/types';

interface TableContextInfo {
  hasMultiCellSelection?: boolean;
  canSplitCell?: boolean;
}

interface ContextMenuState {
  isOpen: boolean;
  position: { x: number; y: number };
  hasSelection: boolean;
  cursorInTable: boolean;
  tableContext: TableContextInfo | null;
}

/**
 * Owns the right-click context-menu surfaces:
 *  - text context menu (cut/copy/paste/pasteAsPlainText/delete/selectAll
 *    + add-comment when there's a selection + table ops when in a cell)
 *  - image context menu (wrap-type swatches + reused text actions)
 *
 * Clipboard and select-all shortcuts come from the command store, so a
 * host's remap shows here; delete keeps its translated key name.
 *
 * The text menu's `addComment` branch needs to mutate comment-management
 * state (selection range, Y position, sidebar visibility, isAddingComment,
 * floatingCommentBtn). To keep this hook independent of comment state
 * ownership, the parent passes a single `onAddComment({ from, to, yPos })`
 * callback that fans out to those setters.
 */
export function useContextMenus({
  commands,
  pagedEditorRef,
  focusActiveEditor,
  runTableAction,
  editorContentRef,
  displayListQueries,
  interactionPageHostRef,
  i18n,
  partEditOpen,
  readOnly = false,
  onAddComment,
}: {
  commands: DocxCommandStore;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  focusActiveEditor: () => void;
  runTableAction: (action: DocxTableAction) => void;
  editorContentRef: React.RefObject<HTMLDivElement | null>;
  displayListQueries: DisplayListQueries | null;
  interactionPageHostRef: React.RefObject<HTMLDivElement | null>;
  i18n: Translations | undefined;
  partEditOpen: boolean;
  /** Offers only Copy and Select all, and no image menu. */
  readOnly?: boolean;
  onAddComment: (range: { from: number; to: number; yPos: number | null }) => void;
}) {
  const { t } = useTranslation();
  const [contextMenu, setContextMenu] = useState<ContextMenuState>({
    isOpen: false,
    position: { x: 0, y: 0 },
    hasSelection: false,
    cursorInTable: false,
    tableContext: null,
  });

  const imageContextMenu = useImageContextMenu();

  const tableContext = useCallback((): TableContextInfo | null => {
    const session = pagedEditorRef.current?.getYrsSession();
    const target = session ? currentYrsTableTarget(session) : null;
    if (!target) return null;
    return {
      canSplitCell: true,
      hasMultiCellSelection:
        target.range.anchor.row !== target.range.head.row ||
        target.range.anchor.column !== target.range.head.column,
    };
  }, [pagedEditorRef]);

  // The body editor's right-click is wired through PagedEditor's
  // onContextMenu (handleContextMenu below). This handler is mounted on the
  // outer editor shell to catch HF-region clicks while the inline editor is
  // open — the body's plumbing won't fire for HF clicks.
  const handleEditorContextMenu = useCallback(
    (e: React.MouseEvent) => {
      const target = e.target as HTMLElement | null;
      // A body/page right-click is handled by PagedEditor's own onContextMenu;
      // this shell-level handler only fires the HF-fallback for clicks OUTSIDE
      // the page area. Accept both renderers' page hosts (`.paged-editor__pages`
      // painter, `.canvas-pages` canvas) so canvas body right-clicks don't
      // mis-route into the HF context path.
      if (isWithinPageArea(target) && !target?.closest('.hf-inline-editor')) {
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      const currentTable = tableContext();
      const selection = pagedEditorRef.current?.getSelectionRange();
      setContextMenu({
        isOpen: true,
        position: { x: e.clientX, y: e.clientY },
        hasSelection: !!selection && selection.from !== selection.to,
        cursorInTable: currentTable != null,
        tableContext: currentTable,
      });
    },
    [pagedEditorRef, tableContext]
  );

  const handleContextMenu = useCallback(
    (data: {
      x: number;
      y: number;
      hasSelection: boolean;
      image?: {
        pos: number;
        wrapType: WrapType;
        cssFloat?: 'left' | 'right' | 'none' | null;
        inlinePositionEmu?: { horizontalEmu: number; verticalEmu: number };
      } | null;
    }) => {
      // An image right-click takes priority over the text context menu.
      if (data.image && !readOnly) {
        imageContextMenu.openForImage({
          x: data.x,
          y: data.y,
          wrapType: data.image.wrapType,
          cssFloat: data.image.cssFloat,
          pos: data.image.pos,
          inlinePositionEmu: data.image.inlinePositionEmu,
        });
        return;
      }
      const currentTable = tableContext();
      setContextMenu({
        isOpen: true,
        position: data,
        hasSelection: data.hasSelection,
        cursorInTable: currentTable != null,
        tableContext: currentTable,
      });
    },
    [imageContextMenu, readOnly, tableContext]
  );

  const handleImageWrapApply = useCallback(
    (target: ImageLayoutTarget) => {
      if (imageContextMenu.imagePos === null) return;
      // For inline → anchor, hand the captured EMU offset to the command so
      // the new float lands where the inline glyph used to sit.
      const opts = imageContextMenu.inlinePositionEmu
        ? { initialPositionEmu: imageContextMenu.inlinePositionEmu }
        : undefined;
      const paged = pagedEditorRef.current;
      paged?.applyYrsCommand({
        type: 'imageWrap',
        pmPos: imageContextMenu.imagePos,
        target,
        options: opts,
      });
    },
    [
      imageContextMenu.imagePos,
      imageContextMenu.inlinePositionEmu,
      pagedEditorRef,
    ]
  );

  const shortcutRevision = useDocxShortcutRevision(commands);
  const keys = useMemo(() => {
    void shortcutRevision;
    return (id: DocxCommandId) => ({
      shortcut: storeShortcutLabel(commands, id, t) ?? undefined,
      ariaShortcut: storeAriaShortcut(commands, id) ?? undefined,
    });
  }, [commands, shortcutRevision, t]);

  // Cut / Copy / Paste / Delete ride along inside the image context menu so
  // users don't need to flip menus to do basic clipboard work on the
  // selected image.
  const imageContextMenuTextActions = useMemo(
    () => [
      {
        action: 'cut' as TextContextAction,
        label: t('contextMenu.cut'),
        ...keys('cut'),
      },
      {
        action: 'copy' as TextContextAction,
        label: t('contextMenu.copy'),
        ...keys('copy'),
      },
      {
        action: 'paste' as TextContextAction,
        label: t('contextMenu.paste'),
        ...keys('paste'),
        dividerAfter: true,
      },
      {
        action: 'delete' as TextContextAction,
        label: t('contextMenu.delete'),
        shortcut: formatKeys(t('contextMenu.deleteShortcut')),
      },
    ],
    [t, keys]
  );

  const handleContextMenuClose = useCallback(() => {
    setContextMenu({
      isOpen: false,
      position: { x: 0, y: 0 },
      hasSelection: false,
      cursorInTable: false,
      tableContext: null,
    });
  }, []);

  const contextMenuItems = useMemo((): TextContextMenuItem[] => {
    if (readOnly) {
      return [
        {
          action: 'copy',
          label: t('contextMenu.copy'),
          ...keys('copy'),
          dividerAfter: true,
        },
        {
          action: 'selectAll',
          label: t('contextMenu.selectAll'),
          ...keys('selectAll'),
        },
      ];
    }
    const items: TextContextMenuItem[] = [
      {
        action: 'cut',
        label: t('contextMenu.cut'),
        ...keys('cut'),
      },
      {
        action: 'copy',
        label: t('contextMenu.copy'),
        ...keys('copy'),
      },
      {
        action: 'paste',
        label: t('contextMenu.paste'),
        ...keys('paste'),
      },
      {
        action: 'pasteAsPlainText',
        label: t('contextMenu.pastePlainText'),
        ...keys('pasteAsPlainText'),
        dividerAfter: true,
      },
      {
        action: 'delete',
        label: t('contextMenu.delete'),
        shortcut: formatKeys(t('contextMenu.deleteShortcut')),
        dividerAfter: !contextMenu.hasSelection && !contextMenu.cursorInTable,
      },
    ];
    if (contextMenu.hasSelection && !partEditOpen) {
      items.push({
        action: 'addComment',
        label: t('contextMenu.addComment'),
        dividerAfter: !contextMenu.cursorInTable,
      });
    }
    if (contextMenu.cursorInTable) {
      items.push(
        { action: 'addRowAbove', label: t('table.insertRowAbove') },
        { action: 'addRowBelow', label: t('table.insertRowBelow') },
        { action: 'deleteRow', label: t('table.deleteRow'), dividerAfter: true },
        { action: 'addColumnLeft', label: t('table.insertColumnLeft') },
        { action: 'addColumnRight', label: t('table.insertColumnRight') },
        { action: 'deleteColumn', label: t('table.deleteColumn') },
        {
          action: 'mergeCells',
          label: i18n?.table?.mergeCells ?? defaultLocale.table.mergeCells,
          disabled: !contextMenu.tableContext?.hasMultiCellSelection,
        },
        {
          action: 'splitCell',
          label: i18n?.table?.splitCell ?? defaultLocale.table.splitCell,
          disabled: !contextMenu.tableContext?.canSplitCell,
          dividerAfter: true,
        },
        {
          action: 'selectTable',
          label: i18n?.table?.selectTable ?? defaultLocale.table.selectTable,
        },
        {
          action: 'deleteTable',
          label: i18n?.table?.deleteTable ?? defaultLocale.table.deleteTable,
          dividerAfter: true,
        }
      );
    }
    items.push({
      action: 'selectAll',
      label: t('contextMenu.selectAll'),
      ...keys('selectAll'),
    });
    return items;
  }, [
    keys,
    contextMenu.hasSelection,
    contextMenu.cursorInTable,
    contextMenu.tableContext,
    i18n,
    partEditOpen,
    readOnly,
    t,
  ]);

  const handleContextMenuAction = useCallback(
    async (action: TextContextAction) => {
      focusActiveEditor();
      const paged = pagedEditorRef.current;
      if (!paged) return;

      switch (action) {
        case 'cut': {
          paged.copySelection(true);
          break;
        }
        case 'copy': {
          const read = paged.readSelectedText();
          if (!read) {
            paged.copySelection(false);
            break;
          }
          const text = await read.catch(() => '');
          if (text) await navigator.clipboard.writeText(text).catch(() => undefined);
          break;
        }
        case 'paste': {
          try {
            const text = await navigator.clipboard.readText();
            const fragment = await readClipboardFragment();
            if (text || fragment) paged.pasteContent(text, fragment);
          } catch {
            // Clipboard access denied.
          }
          break;
        }
        case 'pasteAsPlainText':
          try {
            const text = await navigator.clipboard.readText();
            if (text) paged.insertText(text);
          } catch {
            // Clipboard access denied
          }
          break;
        case 'delete': {
          paged.deleteSelection();
          break;
        }
        case 'selectAll':
          paged.selectAll();
          break;
        case 'addRowAbove':
        case 'addRowBelow':
        case 'deleteRow':
        case 'addColumnLeft':
        case 'addColumnRight':
        case 'deleteColumn':
        case 'mergeCells':
        case 'splitCell':
        case 'selectTable':
        case 'deleteTable':
          runTableAction(action);
          break;
        case 'addComment': {
          if (partEditOpen) break;
          const selection = paged.getSelectionRange();
          if (!selection || selection.from === selection.to) break;
          const { from, to } = selection;
          const anchor = displayListQueries?.anchorRect(from) ?? null;
          const host = interactionPageHostRef.current;
          const target = editorContentRef.current;
          const pageRect =
            anchor && host && displayListQueries
              ? resolveDisplayPageClientRect(host, displayListQueries, anchor.pageIndex)
              : null;
          const pageSize =
            anchor && displayListQueries ? displayListQueries.pageSize(anchor.pageIndex) : null;
          const targetRect = target?.getBoundingClientRect();
          const yPos =
            anchor && pageRect && pageSize && target && targetRect
              ? (pageRect.top -
                  targetRect.top +
                  anchor.y * (pageSize.height > 0 ? pageRect.height / pageSize.height : 1)) /
                effectiveZoom(target)
              : null;
          onAddComment({ from, to, yPos });
          break;
        }
      }
      // TextContextMenu closes itself before running the action.
    },
    [
      focusActiveEditor,
      pagedEditorRef,
      runTableAction,
      editorContentRef,
      displayListQueries,
      interactionPageHostRef,
      onAddComment,
      partEditOpen,
    ]
  );

  return {
    contextMenu,
    imageContextMenu,
    handleEditorContextMenu,
    handleContextMenu,
    handleContextMenuClose,
    handleImageWrapApply,
    imageContextMenuTextActions,
    contextMenuItems,
    handleContextMenuAction,
  };
}
