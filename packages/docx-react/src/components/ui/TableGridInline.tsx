import { useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, ReactElement } from 'react';

interface TableGridInlineProps {
  onInsert: (rows: number, columns: number) => void;
  gridRows?: number;
  gridColumns?: number;
}

const CELL_SIZE = 18;
const CELL_GAP = 2;

const cellStyle: CSSProperties = {
  width: CELL_SIZE,
  height: CELL_SIZE,
  backgroundColor: 'var(--doc-surface)',
  border: '1px solid var(--doc-border)',
  borderRadius: 2,
  transition: 'background-color 0.1s, border-color 0.1s',
  cursor: 'pointer',
};

const cellSelectedStyle: CSSProperties = {
  ...cellStyle,
  backgroundColor: 'var(--doc-primary)',
  border: '1px solid var(--doc-primary)',
};

const labelStyle: CSSProperties = {
  marginTop: 6,
  fontSize: 11,
  fontWeight: 500,
  color: 'var(--doc-text)',
  textAlign: 'center',
};

export function TableGridInline({ onInsert, gridRows = 6, gridColumns = 6 }: TableGridInlineProps) {
  const [selection, setSelection] = useState({ row: 0, col: 0 });
  const [focusedCell, setFocusedCell] = useState({ row: 1, col: 1 });
  const gridRef = useRef<HTMLDivElement>(null);

  const handleKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    const cell = (event.target as HTMLElement).closest<HTMLElement>('[role="gridcell"]');
    if (!cell) return;
    let row = Number(cell.dataset.row);
    let col = Number(cell.dataset.column);
    switch (event.key) {
      case 'ArrowUp':
        row = Math.max(1, row - 1);
        break;
      case 'ArrowDown':
        row = Math.min(gridRows, row + 1);
        break;
      case 'ArrowLeft':
        col = Math.max(1, col - 1);
        break;
      case 'ArrowRight':
        col = Math.min(gridColumns, col + 1);
        break;
      case 'Home':
        col = 1;
        if (event.ctrlKey) row = 1;
        break;
      case 'End':
        col = gridColumns;
        if (event.ctrlKey) row = gridRows;
        break;
      case 'Enter':
      case ' ':
        event.preventDefault();
        event.stopPropagation();
        onInsert(row, col);
        return;
      default:
        return;
    }
    event.preventDefault();
    event.stopPropagation();
    setSelection({ row, col });
    gridRef.current
      ?.querySelector<HTMLElement>(`[data-row="${row}"][data-column="${col}"]`)
      ?.focus();
  };

  const gridRowsContent: ReactElement[] = [];
  for (let row = 1; row <= gridRows; row++) {
    const cells: ReactElement[] = [];
    for (let col = 1; col <= gridColumns; col++) {
      const isSelected = row <= selection.row && col <= selection.col;
      const isFocused = focusedCell.row === row && focusedCell.col === col;
      cells.push(
        <div
          key={col}
          style={isSelected ? cellSelectedStyle : cellStyle}
          onMouseEnter={() => setSelection({ row, col })}
          onFocus={() => {
            setFocusedCell({ row, col });
            setSelection({ row, col });
          }}
          onClick={() => onInsert(row, col)}
          role="gridcell"
          aria-label={`${col} columns, ${row} rows`}
          aria-selected={isSelected}
          data-row={row}
          data-column={col}
          tabIndex={isFocused ? 0 : -1}
        />
      );
    }
    gridRowsContent.push(
      <div key={row} role="row" style={{ display: 'flex', gap: CELL_GAP }}>
        {cells}
      </div>
    );
  }

  const gridLabel =
    selection.row > 0 && selection.col > 0 ? `${selection.col} × ${selection.row}` : 'Select size';

  return (
    <div>
      <div
        ref={gridRef}
        onKeyDown={handleKeyDown}
        style={{
          display: 'flex',
          flexDirection: 'column',
          gap: CELL_GAP,
        }}
        onMouseLeave={() => {
          setSelection(
            gridRef.current?.contains(document.activeElement) ? focusedCell : { row: 0, col: 0 }
          );
        }}
        role="grid"
        aria-label="Table size selector"
        aria-multiselectable="true"
      >
        {gridRowsContent}
      </div>
      <div style={labelStyle} aria-live="polite">
        {gridLabel}
      </div>
    </div>
  );
}
