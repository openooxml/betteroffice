import { useMemo, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent } from 'react';
import type { TFunction } from '@betteroffice/vsdx-i18n';
import type { StandardShape } from './shapeLibrary';

export interface ShapesPanelProps {
  shapes: readonly StandardShape[];
  collapsed: boolean;
  onToggleCollapsed: () => void;
  onInsert: (shape: StandardShape) => void;
  t: TFunction;
  className?: string;
}

const styles: Record<string, CSSProperties> = {
  root: { display: 'flex', minWidth: 0, height: '100%', background: '#fff', color: '#242424', font: '400 13px ui-sans-serif, system-ui, sans-serif', borderRight: '1px solid #e0e0e0', boxSizing: 'border-box' },
  rail: { display: 'flex', flexDirection: 'column', alignItems: 'center', width: 44, padding: '8px 6px', background: '#f7f7f7', borderRight: '1px solid #e5e5e5', boxSizing: 'border-box' },
  railButton: { appearance: 'none', display: 'grid', placeItems: 'center', width: 30, height: 30, padding: 0, border: 0, borderRadius: 4, background: '#dbeafe', color: '#0f6cbd', cursor: 'pointer' },
  content: { display: 'flex', flexDirection: 'column', minWidth: 220, width: 280, height: '100%' },
  header: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', minHeight: 44, padding: '0 10px 0 14px', borderBottom: '1px solid #e5e5e5', fontWeight: 600, fontSize: 14 },
  toggle: { appearance: 'none', display: 'grid', placeItems: 'center', width: 28, height: 28, padding: 0, border: 0, borderRadius: 4, background: 'transparent', color: '#424242', cursor: 'pointer' },
  search: { width: 'calc(100% - 24px)', height: 32, margin: 12, padding: '0 9px', border: '1px solid #bdbdbd', borderRadius: 3, outline: 0, color: '#242424', font: '400 13px ui-sans-serif, system-ui, sans-serif', boxSizing: 'border-box' },
  heading: { margin: '3px 12px 10px', color: '#424242', fontWeight: 600, fontSize: 12 },
  grid: { display: 'flex', flexDirection: 'column', gap: 5, padding: '0 10px 12px', overflowY: 'auto' },
  row: { display: 'grid', gridTemplateColumns: 'repeat(3, minmax(0, 1fr))', gap: 5 },
  cell: { display: 'flex', minWidth: 0 },
  tile: { appearance: 'none', flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', minWidth: 0, minHeight: 92, padding: '7px 3px 5px', border: '1px solid transparent', borderRadius: 3, background: 'transparent', color: '#242424', cursor: 'pointer', font: '400 11px ui-sans-serif, system-ui, sans-serif', textAlign: 'center' },
  preview: { width: 54, height: 46, marginBottom: 5, overflow: 'visible', fill: '#fff', stroke: '#424242', strokeWidth: 1.5, vectorEffect: 'non-scaling-stroke' },
  empty: { margin: '20px 12px', color: '#616161', textAlign: 'center' },
};

const COLUMNS = 3;

function nextFocusIndex(key: string, count: number, index: number): number {
  if (key === 'ArrowRight') return (index + 1) % count;
  if (key === 'ArrowLeft') return (index - 1 + count) % count;
  if (key === 'ArrowDown') return (index + COLUMNS) % count;
  if (key === 'ArrowUp') return (index - COLUMNS + count) % count;
  if (key === 'Home') return 0;
  if (key === 'End') return count - 1;
  return index;
}

export function ShapesPanel({ shapes, collapsed, onToggleCollapsed, onInsert, t, className }: ShapesPanelProps) {
  const [query, setQuery] = useState('');
  const [focusIndex, setFocusIndex] = useState(0);
  const tileRefs = useRef(new Map<number, HTMLButtonElement>());
  const filteredShapes = useMemo(() => {
    const normalized = query.trim().toLocaleLowerCase();
    return normalized ? shapes.filter((shape) => t(shape.nameKey).toLocaleLowerCase().includes(normalized)) : shapes;
  }, [query, shapes, t]);
  const activeIndex = filteredShapes.length ? Math.min(focusIndex, filteredShapes.length - 1) : 0;
  const moveFocus = (event: KeyboardEvent<HTMLButtonElement>, index: number) => {
    const next = nextFocusIndex(event.key, filteredShapes.length, index);
    if (next === index) return;
    event.preventDefault();
    setFocusIndex(next);
    tileRefs.current.get(next)?.focus();
  };
  const rows = Array.from({ length: Math.ceil(filteredShapes.length / COLUMNS) }, (_, row) => filteredShapes.slice(row * COLUMNS, row * COLUMNS + COLUMNS));
  return (
    <aside className={className} style={styles.root} aria-label={t('shapesPanel.title')}>
      <nav style={styles.rail} aria-label={t('shapesPanel.categoriesLabel')}>
        <ul style={{ display: 'contents', margin: 0, padding: 0, listStyle: 'none' }}>
          <li>
            <span role="img" aria-label={t('shapesPanel.standardShapes')} aria-current="true" title={t('shapesPanel.standardShapes')} style={styles.railButton}>
              <svg aria-hidden="true" width="18" height="18" viewBox="0 0 18 18"><rect x="3" y="3" width="5" height="5" fill="none" stroke="currentColor" /><circle cx="13" cy="5.5" r="2.5" fill="none" stroke="currentColor" /><path d="M 3 14 L 6 10 L 9 14 Z" fill="none" stroke="currentColor" /></svg>
            </span>
          </li>
          {collapsed && <li><button type="button" aria-label={t('shapesPanel.expand')} aria-expanded="false" title={t('shapesPanel.expand')} onClick={onToggleCollapsed} style={{ ...styles.toggle, marginTop: 8 }}>›</button></li>}
        </ul>
      </nav>
      {!collapsed && (
        <section style={styles.content} aria-label={t('shapesPanel.title')}>
          <header style={styles.header}>
            <span>{t('shapesPanel.title')}</span>
            <button type="button" aria-label={t('shapesPanel.collapse')} aria-expanded="true" title={t('shapesPanel.collapse')} onClick={onToggleCollapsed} style={styles.toggle}>‹</button>
          </header>
          <input type="search" value={query} onChange={(event) => { setQuery(event.target.value); setFocusIndex(0); }} placeholder={t('shapesPanel.searchPlaceholder')} aria-label={t('shapesPanel.searchLabel')} style={styles.search} />
          <h2 style={styles.heading}>{t('shapesPanel.standardShapes')}</h2>
          {filteredShapes.length === 0 ? <p style={styles.empty}>{t('shapesPanel.empty')}</p> : (
            <div role="grid" aria-label={t('shapesPanel.standardShapes')} style={styles.grid}>
              {rows.map((row, rowIndex) => (
                <div key={rowIndex} role="row" style={styles.row}>
                  {row.map((shape, columnIndex) => {
                    const index = rowIndex * COLUMNS + columnIndex;
                    return (
                      <div key={shape.id} role="gridcell" style={styles.cell}>
                        <button
                          ref={(element) => { if (element) tileRefs.current.set(index, element); else tileRefs.current.delete(index); }}
                          type="button"
                          tabIndex={index === activeIndex ? 0 : -1}
                          aria-label={t(shape.nameKey)}
                          onFocus={() => setFocusIndex(index)}
                          onClick={() => onInsert(shape)}
                          onKeyDown={(event) => moveFocus(event, index)}
                          style={styles.tile}
                        >
                          <svg aria-hidden="true" viewBox="0 0 1 1" preserveAspectRatio="xMidYMid meet" style={styles.preview}><path d={shape.preview} /></svg>
                          <span>{t(shape.nameKey)}</span>
                        </button>
                      </div>
                    );
                  })}
                </div>
              ))}
            </div>
          )}
        </section>
      )}
    </aside>
  );
}
