import { useCallback, useState } from 'react';
import type { YrsCellBorders } from '@betteroffice/docx/yrs';

import type { TableAction } from '../../ui/TableToolbar';
import { getBuiltinTableStyle } from '../../ui/TableStyleGallery';
import type { PagedEditorRef } from '../PagedEditor';
import { currentYrsSplitCellConfig, currentYrsTableProperties } from '../yrsCommands';
import type { TableProperties } from '../../dialogs/TablePropertiesDialog';

interface SplitCellDialogState {
  isOpen: boolean;
  initialRows: number;
  initialCols: number;
  minRows: number;
  minCols: number;
}

interface BorderSpec {
  style: string;
  size: number;
  color: { rgb: string };
}

/** Table toolbar/dialog routing for the authoritative yrs session. */
export function useTableDialogs({
  pagedEditorRef,
  borderSpecRef,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  borderSpecRef: React.RefObject<BorderSpec>;
}) {
  const [tablePropsOpen, setTablePropsOpen] = useState(false);
  const [splitCellDialogState, setSplitCellDialogState] = useState<SplitCellDialogState>({
    isOpen: false,
    initialRows: 1,
    initialCols: 2,
    minRows: 1,
    minCols: 1,
  });

  const openSplitCellDialog = useCallback((_legacyView?: unknown) => {
    const session = pagedEditorRef.current?.getYrsSession();
    const config = session ? currentYrsSplitCellConfig(session) : null;
    if (!config) return;
    setSplitCellDialogState({ ...config, isOpen: true });
  }, [pagedEditorRef]);

  const currentTableProperties = (() => {
    const session = pagedEditorRef.current?.getYrsSession();
    try {
      return session ? currentYrsTableProperties(session) : undefined;
    } catch {
      return undefined;
    }
  })();

  const handleTablePropertiesApply = useCallback(
    (properties: TableProperties) => {
      pagedEditorRef.current?.applyYrsCommand({ type: 'tableProperties', properties });
    },
    [pagedEditorRef]
  );

  const applyBorders = useCallback(
    (borders: YrsCellBorders) => {
      pagedEditorRef.current?.applyYrsCommand({ type: 'tableSetBorders', borders });
    },
    [pagedEditorRef]
  );

  const allBorders = useCallback(
    (border: BorderSpec): YrsCellBorders => ({
      top: border,
      bottom: border,
      left: border,
      right: border,
      insideH: border,
      insideV: border,
    }),
    []
  );

  const handleTableAction = useCallback(
    (action: TableAction) => {
      const editor = pagedEditorRef.current;
      if (!editor) return;
      const apply = editor.applyYrsCommand;
      if (typeof action === 'object') {
        if (action.type === 'cellFillColor') {
          apply({ type: 'tableCellShading', color: action.color });
        } else if (action.type === 'borderColor') {
          borderSpecRef.current = {
            ...borderSpecRef.current,
            color: { rgb: action.color.replace(/^#/, '') },
          };
          applyBorders(allBorders(borderSpecRef.current));
        } else if (action.type === 'borderWidth') {
          borderSpecRef.current = { ...borderSpecRef.current, size: action.size };
          applyBorders(allBorders(borderSpecRef.current));
        } else if (action.type === 'cellBorder') {
          applyBorders({
            [action.side]: {
              style: action.style,
              size: action.size,
              color: { rgb: action.color.replace(/^#/, '') },
            },
          });
        } else if (action.type === 'openTableProperties') {
          setTablePropsOpen(true);
        } else if (action.type === 'applyTableStyle') {
          const preset = getBuiltinTableStyle(action.styleId);
          if (preset?.tableBorders) applyBorders(preset.tableBorders);
        }
        return;
      }

      switch (action) {
        case 'addRowAbove':
          apply({ type: 'tableInsertRow', side: 'above' });
          break;
        case 'addRowBelow':
          apply({ type: 'tableInsertRow', side: 'below' });
          break;
        case 'addColumnLeft':
          apply({ type: 'tableInsertColumn', side: 'left' });
          break;
        case 'addColumnRight':
          apply({ type: 'tableInsertColumn', side: 'right' });
          break;
        case 'deleteRow':
          apply({ type: 'tableDeleteRow' });
          break;
        case 'deleteColumn':
          apply({ type: 'tableDeleteColumn' });
          break;
        case 'deleteTable':
          apply({ type: 'tableDelete' });
          break;
        case 'mergeCells':
          apply({ type: 'tableMergeCells' });
          break;
        case 'splitCell':
          openSplitCellDialog();
          break;
        case 'selectTable':
          apply({ type: 'tableSelect', target: 'table' });
          break;
        case 'selectRow':
          apply({ type: 'tableSelect', target: 'row' });
          break;
        case 'selectColumn':
          apply({ type: 'tableSelect', target: 'column' });
          break;
        case 'borderAll':
          applyBorders(allBorders(borderSpecRef.current));
          break;
        case 'borderOutside':
          applyBorders({
            top: borderSpecRef.current,
            bottom: borderSpecRef.current,
            left: borderSpecRef.current,
            right: borderSpecRef.current,
          });
          break;
        case 'borderInside':
          applyBorders({ insideH: borderSpecRef.current, insideV: borderSpecRef.current });
          break;
        case 'borderNone':
          applyBorders(allBorders({ style: 'none', size: 0, color: { rgb: '000000' } }));
          break;
        case 'borderTop':
        case 'borderBottom':
        case 'borderLeft':
        case 'borderRight':
          applyBorders({
            [action.slice('border'.length).toLowerCase()]: borderSpecRef.current,
          });
          break;
      }
    },
    [allBorders, applyBorders, borderSpecRef, openSplitCellDialog, pagedEditorRef]
  );

  const handleSplitCellDialogClose = useCallback(() => {
    setSplitCellDialogState((previous) => ({ ...previous, isOpen: false }));
  }, []);

  const handleSplitCellDialogApply = useCallback(
    (rows: number, columns: number) => {
      pagedEditorRef.current?.applyYrsCommand({ type: 'tableSplitCell', rows, columns });
      setSplitCellDialogState((previous) => ({ ...previous, isOpen: false }));
      pagedEditorRef.current?.focus();
    },
    [pagedEditorRef]
  );

  return {
    tablePropsOpen,
    setTablePropsOpen,
    currentTableProperties,
    handleTablePropertiesApply,
    splitCellDialogState,
    openSplitCellDialog,
    handleTableAction,
    handleSplitCellDialogClose,
    handleSplitCellDialogApply,
  };
}
