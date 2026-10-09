import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, ReactNode } from 'react';
import type { TFunction } from '@betteroffice/vsdx-i18n';
import type { CellSnapshot, DiagramSnapshot, ShapeSnapshot } from '@betteroffice/vsdx';
import type { VsdxShapeSelection } from '../../VsdxEditor';

export const MAX_EXPLORER_DEPTH = 64;

export interface DrawingExplorerProps {
  snapshot: DiagramSnapshot | null;
  activePageIndex: number;
  selection: VsdxShapeSelection | null;
  onSelectPage: (index: number) => void;
  onSelectShape: (selection: VsdxShapeSelection | null) => void;
  collapsed: boolean;
  onToggleCollapsed: () => void;
  t: TFunction;
  className?: string;
}

export interface ExplorerSection {
  key: string;
  name: string;
  sectionIndex: number | null;
  cells: CellSnapshot[];
}

export interface ExplorerRow {
  key: string;
  heading: string | null;
  cells: CellSnapshot[];
}

export function isOneDimensional(shape: ShapeSnapshot): boolean {
  const cell = shape.cells.find((item) => item.locator.section === null && item.locator.row === null && item.locator.cellName === 'OneD');
  const text = (cell?.value ?? cell?.formula ?? '').trim();
  return text !== '' && Number(text) === 1;
}

export function findShapePath(shapes: readonly ShapeSnapshot[], shapeId: string, depth = 0): ShapeSnapshot[] | null {
  if (depth > MAX_EXPLORER_DEPTH) return null;
  for (const shape of shapes) {
    if (shape.id === shapeId) return [shape];
    const nested = findShapePath(shape.children, shapeId, depth + 1);
    if (nested) return [shape, ...nested];
  }
  return null;
}

export function groupShapeSections(shape: ShapeSnapshot): { shapeCells: CellSnapshot[]; sections: ExplorerSection[] } {
  const shapeCells: CellSnapshot[] = [];
  const sections: ExplorerSection[] = [];
  const byKey = new Map<string, ExplorerSection>();
  for (const cell of shape.cells) {
    const section = cell.locator.section;
    if (section === null) {
      shapeCells.push(cell);
      continue;
    }
    const sectionIndex = cell.locator.sectionIndex ?? null;
    const key = `${section} ${sectionIndex ?? ''}`;
    let group = byKey.get(key);
    if (!group) {
      group = { key, name: section, sectionIndex, cells: [] };
      byKey.set(key, group);
      sections.push(group);
    }
    group.cells.push(cell);
  }
  return { shapeCells, sections };
}

export function groupSectionRows(cells: readonly CellSnapshot[], rowLabel: (index: number) => string): ExplorerRow[] {
  const rows: ExplorerRow[] = [];
  const byKey = new Map<string, ExplorerRow>();
  for (const cell of cells) {
    const row = cell.locator.row;
    const key = row === null ? '' : 'index' in row ? `index:${row.index}` : `name:${row.name}`;
    let group = byKey.get(key);
    if (!group) {
      group = { key, heading: rowHeading(row, cell.rowType, rowLabel), cells: [] };
      byKey.set(key, group);
      rows.push(group);
    }
    group.cells.push(cell);
  }
  return rows;
}

function rowHeading(row: { index: number } | { name: string } | null, rowType: string | undefined, rowLabel: (index: number) => string): string | null {
  if (row === null) return null;
  if ('index' in row) {
    const base = rowLabel(row.index);
    return rowType ? `${base} (${rowType})` : base;
  }
  return rowType && rowType !== row.name ? `${row.name} (${rowType})` : row.name;
}

function shapeLabel(shape: ShapeSnapshot, t: TFunction): string {
  const base = t('explorer.shape', { id: shape.sourceId });
  const named = shape.name ? `${base} "${shape.name}"` : base;
  const tags: string[] = [];
  if (shape.children.length > 0) tags.push(t('explorer.groupTag'));
  if (isOneDimensional(shape)) tags.push(t('explorer.oneDTag'));
  return tags.length > 0 ? `${named} (${tags.join(', ')})` : named;
}

function sectionLabel(section: ExplorerSection): string {
  return section.sectionIndex !== null && section.sectionIndex > 0 ? `${section.name} ${section.sectionIndex + 1}` : section.name;
}

function pageIdsKey(snapshot: DiagramSnapshot | null): string {
  return snapshot ? snapshot.pages.map((page) => page.id).join(' ') : '';
}

export function DrawingExplorer({ snapshot, activePageIndex, selection, onSelectPage, onSelectShape, collapsed, onToggleCollapsed, t, className }: DrawingExplorerProps) {
  const [expanded, setExpanded] = useState<ReadonlySet<string>>(() => new Set(['document', 'pages']));
  const expandedPageRef = useRef<string | null>(null);
  const revealedRef = useRef<string | null>(null);
  const documentKey = pageIdsKey(snapshot);

  useLayoutEffect(() => {
    const activeIndex = Math.max(0, Math.min(activePageIndex, (snapshot?.pages.length ?? 1) - 1));
    const expandedPageKey = `${documentKey}:${activeIndex}`;
    if (expandedPageRef.current === expandedPageKey) return;
    expandedPageRef.current = expandedPageKey;
    const active = snapshot?.pages[activeIndex];
    const next = new Set<string>(['document', 'pages']);
    if (active) {
      next.add(`page:${active.id}`);
      next.add(`shapes:${active.id}`);
    }
    setExpanded(next);
  }, [activePageIndex, documentKey, snapshot]);

  useEffect(() => {
    if (!snapshot || !selection) { revealedRef.current = null; return; }
    const revealKey = `${selection.pageId} ${selection.shapeId}`;
    if (revealedRef.current === revealKey) return;
    const page = snapshot.pages.find((entry) => entry.id === selection.pageId);
    if (!page) return;
    const path = findShapePath(page.shapes, selection.shapeId);
    if (!path) return;
    revealedRef.current = revealKey;
    setExpanded((previous) => {
      const next = new Set(previous);
      next.add('document');
      next.add('pages');
      next.add(`page:${page.id}`);
      next.add(`shapes:${page.id}`);
      for (const shape of path.slice(0, -1)) next.add(`shape:${page.id}:${shape.id}`);
      return next;
    });
  }, [snapshot, selection]);

  const toggle = (key: string) => {
    setExpanded((previous) => {
      const next = new Set(previous);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const selectShape = (pageIndex: number, pageId: string, shapeId: string) => {
    if (pageIndex !== activePageIndex) onSelectPage(pageIndex);
    onSelectShape({ pageId, shapeId, hit: { kind: 'shape', shapeId } });
  };

  const moveFocus = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
    const buttons = Array.from(event.currentTarget.querySelectorAll<HTMLButtonElement>('button[data-explorer-node]'));
    if (!buttons.length) return;
    const active = document.activeElement as HTMLElement | null;
    const index = active ? buttons.indexOf(active as HTMLButtonElement) : -1;
    const next = index < 0 ? buttons[event.key === 'ArrowDown' ? 0 : buttons.length - 1] : buttons[index + (event.key === 'ArrowDown' ? 1 : -1)];
    if (next) {
      event.preventDefault();
      next.focus();
    }
  };

  if (collapsed) {
    return (
      <aside className={className} style={styles.rail} aria-label={t('explorer.title')}>
        <button type="button" aria-label={t('explorer.expand')} aria-expanded="false" title={t('explorer.expand')} onClick={onToggleCollapsed} style={styles.railButton}>
          <svg aria-hidden="true" width="16" height="16" viewBox="0 0 16 16"><path d="M 10 3 L 5 8 L 10 13" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></svg>
        </button>
      </aside>
    );
  }

  return (
    <aside className={className} style={styles.root} aria-label={t('explorer.title')}>
      <header style={styles.header}>
        <span>{t('explorer.title')}</span>
        <button type="button" aria-label={t('explorer.collapse')} aria-expanded="true" title={t('explorer.collapse')} onClick={onToggleCollapsed} style={styles.toggle}>
          <svg aria-hidden="true" width="16" height="16" viewBox="0 0 16 16"><path d="M 6 3 L 11 8 L 6 13" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></svg>
        </button>
      </header>
      <div role="tree" aria-label={t('explorer.title')} onKeyDown={moveFocus} style={styles.tree}>
        {!snapshot || snapshot.pages.length === 0 ? (
          <p style={styles.empty}>{t('explorer.empty')}</p>
        ) : (
          <TreeNode nodeKey="document" level={1} label={t('explorer.document')} selected={false} expandable expanded={expanded.has('document')} onToggle={() => toggle('document')} onActivate={() => {}}>
            {expanded.has('document') && (
            <TreeNode nodeKey="pages" level={2} label={t('explorer.pages')} selected={false} expandable expanded={expanded.has('pages')} onToggle={() => toggle('pages')} onActivate={() => {}}>
              {expanded.has('pages') && snapshot.pages.map((page, pageIndex) => (
                <TreeNode
                  key={page.id}
                  nodeKey={`page:${page.id}`}
                  level={3}
                  label={page.name ?? t('pages.fallbackTitle', { number: pageIndex + 1 })}
                  selected={pageIndex === activePageIndex}
                  expandable
                  expanded={expanded.has(`page:${page.id}`)}
                  onToggle={() => toggle(`page:${page.id}`)}
                  onActivate={() => onSelectPage(pageIndex)}
                >
                  {expanded.has(`page:${page.id}`) && (
                    <TreeNode nodeKey={`shapes:${page.id}`} level={4} label={t('explorer.shapes')} selected={false} expandable expanded={expanded.has(`shapes:${page.id}`)} onToggle={() => toggle(`shapes:${page.id}`)} onActivate={() => onSelectPage(pageIndex)}>
                      {expanded.has(`shapes:${page.id}`) && (page.shapes.length === 0 ? (
                        <p style={styles.empty}>{t('explorer.emptyShapes')}</p>
                      ) : (
                        page.shapes.map((shape) => (
                          <ShapeNode key={shape.id} page={snapshot.pages[pageIndex]} pageIndex={pageIndex} shape={shape} depth={0} level={5} expanded={expanded} selection={selection} onToggle={toggle} onSelectShape={selectShape} t={t} />
                        ))
                      ))}
                    </TreeNode>
                  )}
                </TreeNode>
              ))}
            </TreeNode>
            )}
          </TreeNode>
        )}
      </div>
    </aside>
  );
}

interface TreeNodeProps {
  nodeKey: string;
  level: number;
  label: string;
  selected: boolean;
  expandable: boolean;
  expanded: boolean;
  onToggle: () => void;
  onActivate: () => void;
  children?: ReactNode;
}

function TreeNode({ nodeKey, level, label, selected, expandable, expanded, onToggle, onActivate, children }: TreeNodeProps) {
  const onKeyDown = (event: KeyboardEvent<HTMLButtonElement>) => {
    if (event.key === 'ArrowRight' && expandable && !expanded) {
      event.preventDefault();
      onToggle();
    } else if (event.key === 'ArrowLeft' && expandable && expanded) {
      event.preventDefault();
      onToggle();
    }
  };
  return (
    <div role="treeitem" aria-level={level} aria-selected={selected} {...(expandable ? { 'aria-expanded': expanded } : {})}>
      <div style={{ ...styles.row, paddingLeft: 8 + level * 10, background: selected ? '#e8f0fe' : 'transparent' }}>
        <span aria-hidden="true" style={styles.chevron}>{expandable ? (expanded ? '▾' : '▸') : ''}</span>
        <button type="button" data-explorer-node data-explorer-key={nodeKey} onClick={() => { onActivate(); if (expandable) onToggle(); }} onKeyDown={onKeyDown} aria-expanded={expandable ? expanded : undefined} title={label} style={{ ...styles.label, fontWeight: selected ? 600 : 400 }}>
          {label}
        </button>
      </div>
      {expandable && expanded && children ? <div role="group">{children}</div> : null}
    </div>
  );
}

interface ShapeNodeProps {
  page: { id: string };
  pageIndex: number;
  shape: ShapeSnapshot;
  depth: number;
  level: number;
  expanded: ReadonlySet<string>;
  selection: VsdxShapeSelection | null;
  onToggle: (key: string) => void;
  onSelectShape: (pageIndex: number, pageId: string, shapeId: string) => void;
  t: TFunction;
}

function ShapeNode({ page, pageIndex, shape, depth, level, expanded, selection, onToggle, onSelectShape, t }: ShapeNodeProps) {
  const nodeKey = `shape:${page.id}:${shape.id}`;
  const isOpen = expanded.has(nodeKey);
  const selected = selection?.pageId === page.id && selection?.shapeId === shape.id;
  if (depth >= MAX_EXPLORER_DEPTH) {
    return (
      <div role="treeitem" aria-level={level} aria-selected={false}>
        <div style={{ ...styles.row, paddingLeft: 8 + level * 10 }}>
          <span aria-hidden="true" style={styles.chevron} />
          <span style={styles.limit}>{t('explorer.depthLimit')}</span>
        </div>
      </div>
    );
  }
  const expandable = shape.children.length > 0 || shape.cells.length > 0;
  const visibleSections: ExplorerSection[] = [];
  if (isOpen) {
    const { shapeCells, sections } = groupShapeSections(shape);
    if (shapeCells.length > 0) visibleSections.push({ key: `${nodeKey}:shape`, name: t('explorer.shapeSection'), sectionIndex: null, cells: shapeCells });
    visibleSections.push(...sections);
  }
  return (
    <TreeNode nodeKey={nodeKey} level={level} label={shapeLabel(shape, t)} selected={selected} expandable={expandable} expanded={isOpen} onToggle={() => onToggle(nodeKey)} onActivate={() => onSelectShape(pageIndex, page.id, shape.id)}>
      {isOpen && shape.children.map((child) => (
        <ShapeNode key={child.id} page={page} pageIndex={pageIndex} shape={child} depth={depth + 1} level={level + 1} expanded={expanded} selection={selection} onToggle={onToggle} onSelectShape={onSelectShape} t={t} />
      ))}
      {isOpen && visibleSections.map((section) => (
        <SectionNode key={section.key} pageId={page.id} shapeId={shape.id} section={section} level={level + 1} expanded={expanded} onToggle={onToggle} t={t} />
      ))}
    </TreeNode>
  );
}

interface SectionNodeProps {
  pageId: string;
  shapeId: string;
  section: ExplorerSection;
  level: number;
  expanded: ReadonlySet<string>;
  onToggle: (key: string) => void;
  t: TFunction;
}

function SectionNode({ pageId, shapeId, section, level, expanded, onToggle, t }: SectionNodeProps) {
  const nodeKey = `section:${pageId}:${shapeId}:${section.key}`;
  const isOpen = expanded.has(nodeKey);
  const rows = isOpen ? groupSectionRows(section.cells, (index) => t('explorer.row', { index: index + 1 })) : [];
  return (
    <TreeNode nodeKey={nodeKey} level={level} label={sectionLabel(section)} selected={false} expandable onToggle={() => onToggle(nodeKey)} expanded={isOpen} onActivate={() => {}}>
      <div style={{ ...styles.cellHeader, paddingLeft: 8 + (level + 1) * 10 }}>
        <span>{t('explorer.cell')}</span>
        <span>{t('explorer.formula')}</span>
        <span>{t('explorer.value')}</span>
      </div>
      {rows.map((row) => (
        <div key={row.key}>
          {row.heading !== null && <div style={{ ...styles.rowHeading, paddingLeft: 8 + (level + 1) * 10 }}>{row.heading}</div>}
          {row.cells.map((cell) => (
            <div key={`${row.key}:${cell.name}`} title={`${cell.name}`} style={{ ...styles.cellRow, paddingLeft: 8 + (level + 1) * 10 }}>
              <span style={styles.cellName}>{cell.name}</span>
              <code style={styles.cellFormula}>{cell.formula ?? '—'}</code>
              <code style={styles.cellValue}>{cell.value ?? '—'}</code>
            </div>
          ))}
        </div>
      ))}
    </TreeNode>
  );
}

const styles: Record<string, CSSProperties> = {
  root: { display: 'flex', flexDirection: 'column', width: 300, minWidth: 300, height: '100%', background: '#fff', color: '#242424', font: '400 13px ui-sans-serif, system-ui, sans-serif', borderLeft: '1px solid #e0e0e0', boxSizing: 'border-box' },
  rail: { display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '8px 0', background: '#fff', borderLeft: '1px solid #e0e0e0', boxSizing: 'border-box' },
  railButton: { appearance: 'none', display: 'grid', placeItems: 'center', width: 28, height: 28, padding: 0, border: 0, borderRadius: 4, background: 'transparent', color: '#424242', cursor: 'pointer' },
  header: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', minHeight: 44, padding: '0 6px 0 14px', borderBottom: '1px solid #e5e5e5', fontWeight: 600, fontSize: 14 },
  toggle: { appearance: 'none', display: 'grid', placeItems: 'center', width: 28, height: 28, padding: 0, border: 0, borderRadius: 4, background: 'transparent', color: '#424242', cursor: 'pointer' },
  tree: { flex: 1, minHeight: 0, overflowY: 'auto', padding: '4px 0 12px' },
  row: { display: 'flex', alignItems: 'center', minHeight: 26, paddingRight: 8, boxSizing: 'border-box' },
  chevron: { flex: '0 0 auto', width: 18, color: '#616161', fontSize: 11, textAlign: 'center' },
  label: { appearance: 'none', flex: 1, minWidth: 0, padding: '3px 6px', overflow: 'hidden', border: 0, borderRadius: 4, background: 'transparent', color: 'inherit', cursor: 'pointer', font: 'inherit', textAlign: 'left', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  limit: { color: '#616161', fontStyle: 'italic' },
  empty: { margin: '12px 14px', color: '#616161' },
  rowHeading: { paddingTop: 4, paddingRight: 8, color: '#424242', fontSize: 11, fontWeight: 600 },
  cellHeader: { display: 'grid', gridTemplateColumns: 'minmax(0, 0.9fr) minmax(0, 1.1fr) minmax(0, 1fr)', gap: 6, paddingTop: 4, paddingRight: 8, color: '#616161', fontSize: 11, fontWeight: 600 },
  cellRow: { display: 'grid', gridTemplateColumns: 'minmax(0, 0.9fr) minmax(0, 1.1fr) minmax(0, 1fr)', gap: 6, paddingTop: 1, paddingBottom: 1, paddingRight: 8, fontSize: 12 },
  cellName: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  cellFormula: { overflow: 'hidden', fontFamily: 'ui-monospace, monospace', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  cellValue: { overflow: 'hidden', fontFamily: 'ui-monospace, monospace', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
};
