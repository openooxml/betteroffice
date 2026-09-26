import { useCallback, useState } from 'react';
import type { YrsCellBorders } from '@betteroffice/docx/yrs';

import type { DocxTableAction } from '../../../commands/types';
import { getBuiltinTableStyle } from '../../ui/TableStyleGallery';
import type { PagedEditorRef } from '../PagedEditor';
import {
  currentYrsSplitCellConfig,
  currentYrsTableProperties,
  type YrsEditorCommand,
} from '../yrsCommands';
import type { TableProperties } from '../../dialogs/TablePropertiesDialog';
import type { DocxTableActionOutcome } from './useDocxCommands';

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

/**
 * Table toolbar/dialog routing for the authoritative yrs session. `apply`
 * runs a command immediately; `complete` finishes a dialog against the table
 * selection it opened for.
 */
export function useTableDialogs({
  pagedEditorRef,
  borderSpecRef,
  apply,
  complete,
}: {
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  borderSpecRef: React.RefObject<BorderSpec>;
  apply: (command: YrsEditorCommand) => boolean;
  complete: (dialog: 'splitCell' | 'tableProperties', command: YrsEditorCommand) => void;
}) {
  const [tablePropsOpen, setTablePropsOpen] = useState(false);
  const [splitCellDialogState, setSplitCellDialogState] = useState<SplitCellDialogState>({
    isOpen: false,
    initialRows: 1,
    initialCols: 2,
    minRows: 1,
    minCols: 1,
  });

  const openSplitCellDialog = useCallback((): boolean => {
    const session = pagedEditorRef.current?.getYrsSession();
    const config = session ? currentYrsSplitCellConfig(session) : null;
    if (!config) return false;
    setSplitCellDialogState({ ...config, isOpen: true });
    return true;
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
      complete('tableProperties', { type: 'tableProperties', properties });
    },
    [complete]
  );

  const applyBorders = useCallback(
    (borders: YrsCellBorders) => apply({ type: 'tableSetBorders', borders }),
    [apply]
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
    (action: DocxTableAction): DocxTableActionOutcome => {
      if (typeof action === 'object') {
        switch (action.type) {
          case 'cellFillColor':
            return apply({ type: 'tableCellShading', color: action.color });
          case 'borderColor':
            borderSpecRef.current = {
              ...borderSpecRef.current,
              color: { rgb: action.color.replace(/^#/, '') },
            };
            return applyBorders(allBorders(borderSpecRef.current));
          case 'borderWidth':
            borderSpecRef.current = { ...borderSpecRef.current, size: action.size };
            return applyBorders(allBorders(borderSpecRef.current));
          case 'cellBorder': {
            const border = {
              style: action.style,
              size: action.size,
              color: { rgb: action.color.replace(/^#/, '') },
            };
            return applyBorders(
              action.side === 'all' ? allBorders(border) : { [action.side]: border }
            );
          }
          case 'tableProperties':
            return apply({ type: 'tableProperties', properties: action.props });
          case 'openTableProperties':
            setTablePropsOpen(true);
            return 'opened';
          case 'applyTableStyle': {
            const preset = getBuiltinTableStyle(action.styleId);
            return preset?.tableBorders ? applyBorders(preset.tableBorders) : false;
          }
          default:
            return false;
        }
      }

      switch (action) {
        case 'addRowAbove':
          return apply({ type: 'tableInsertRow', side: 'above' });
        case 'addRowBelow':
          return apply({ type: 'tableInsertRow', side: 'below' });
        case 'addColumnLeft':
          return apply({ type: 'tableInsertColumn', side: 'left' });
        case 'addColumnRight':
          return apply({ type: 'tableInsertColumn', side: 'right' });
        case 'deleteRow':
          return apply({ type: 'tableDeleteRow' });
        case 'deleteColumn':
          return apply({ type: 'tableDeleteColumn' });
        case 'deleteTable':
          return apply({ type: 'tableDelete' });
        case 'mergeCells':
          return apply({ type: 'tableMergeCells' });
        case 'splitCell':
          return openSplitCellDialog() ? 'opened' : false;
        case 'selectTable':
          return apply({ type: 'tableSelect', target: 'table' });
        case 'selectRow':
          return apply({ type: 'tableSelect', target: 'row' });
        case 'selectColumn':
          return apply({ type: 'tableSelect', target: 'column' });
        case 'borderAll':
          return applyBorders(allBorders(borderSpecRef.current));
        case 'borderOutside':
          return applyBorders({
            top: borderSpecRef.current,
            bottom: borderSpecRef.current,
            left: borderSpecRef.current,
            right: borderSpecRef.current,
          });
        case 'borderInside':
          return applyBorders({ insideH: borderSpecRef.current, insideV: borderSpecRef.current });
        case 'borderNone':
          return applyBorders(allBorders({ style: 'none', size: 0, color: { rgb: '000000' } }));
        case 'borderTop':
        case 'borderBottom':
        case 'borderLeft':
        case 'borderRight':
          return applyBorders({
            [action.slice('border'.length).toLowerCase()]: borderSpecRef.current,
          });
      }
    },
    [allBorders, apply, applyBorders, borderSpecRef, openSplitCellDialog]
  );

  const handleSplitCellDialogClose = useCallback(() => {
    setSplitCellDialogState((previous) => ({ ...previous, isOpen: false }));
  }, []);

  const handleSplitCellDialogApply = useCallback(
    (rows: number, columns: number) => {
      complete('splitCell', { type: 'tableSplitCell', rows, columns });
      setSplitCellDialogState((previous) => ({ ...previous, isOpen: false }));
      pagedEditorRef.current?.focus();
    },
    [complete, pagedEditorRef]
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
