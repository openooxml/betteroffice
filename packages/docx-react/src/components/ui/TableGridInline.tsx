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
  const [hoverRows, setHoverRows] = useState(0);
  const [hoverCols, setHoverCols] = useState(0);
  const [focusedCell, setFocusedCell] = useState({ row: 1, col: 1 });
  const cellRefs = useRef(new Map<string, HTMLDivElement>());

  const gridCells: ReactElement[] = [];
  for (let row = 1; row <= gridRows; row++) {
    for (let col = 1; col <= gridColumns; col++) {
      const isSelected = row <= hoverRows && col <= hoverCols;
      const isFocused = focusedCell.row === row && focusedCell.col === col;
      const key = `${row}-${col}`;
      const handleKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
        const movement: Record<string, [number, number]> = {
          ArrowUp: [-1, 0],
          ArrowDown: [1, 0],
          ArrowLeft: [0, -1],
          ArrowRight: [0, 1],
        };
        const delta = movement[event.key];
        if (delta) {
          event.preventDefault();
          event.stopPropagation();
          const next = {
            row: Math.max(1, Math.min(gridRows, row + delta[0])),
            col: Math.max(1, Math.min(gridColumns, col + delta[1])),
          };
          setFocusedCell(next);
          setHoverRows(next.row);
          setHoverCols(next.col);
          cellRefs.current.get(`${next.row}-${next.col}`)?.focus();
        } else if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          event.stopPropagation();
          onInsert(row, col);
        }
      };

      gridCells.push(
        <div
          key={key}
          ref={(element) => {
            if (element) cellRefs.current.set(key, element);
            else cellRefs.current.delete(key);
          }}
          style={isSelected ? cellSelectedStyle : cellStyle}
          onMouseEnter={() => {
            setHoverRows(row);
            setHoverCols(col);
          }}
          onFocus={() => {
            setFocusedCell({ row, col });
            setHoverRows(row);
            setHoverCols(col);
          }}
          onKeyDown={handleKeyDown}
          onClick={() => onInsert(row, col)}
          role="gridcell"
          aria-label={`${col} columns, ${row} rows`}
          aria-selected={isSelected}
          aria-rowindex={row}
          aria-colindex={col}
          data-row={row}
          data-column={col}
          tabIndex={isFocused ? 0 : -1}
        />
      );
    }
  }

  const gridLabel = hoverRows > 0 && hoverCols > 0 ? `${hoverCols} × ${hoverRows}` : 'Select size';

  return (
    <div>
      <div
        style={{
          display: 'grid',
          gap: CELL_GAP,
          gridTemplateColumns: `repeat(${gridColumns}, ${CELL_SIZE}px)`,
        }}
        onMouseLeave={() => {
          setHoverRows(0);
          setHoverCols(0);
        }}
        role="grid"
        aria-label="Table size selector"
      >
        {gridCells}
      </div>
      <div style={labelStyle} aria-live="polite">
        {gridLabel}
      </div>
    </div>
  );
}
