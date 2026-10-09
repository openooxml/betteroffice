import type { CSSProperties } from 'react';
import type { TFunction } from '@betteroffice/vsdx-i18n';
import type { ValidationIssue } from '@betteroffice/vsdx';

export interface IssuesPanelProps {
  issues: readonly ValidationIssue[];
  collapsed: boolean;
  onToggleCollapsed: () => void;
  onSelect: (issue: ValidationIssue) => void;
  getShapeName: (pageId: string, shapeId: string) => string | null;
  t: TFunction;
  className?: string;
}

const styles: Record<string, CSSProperties> = {
  root: { display: 'flex', flexDirection: 'column', minHeight: 0, maxHeight: '42%', width: '100%', background: '#fff', color: '#242424', font: '400 13px ui-sans-serif, system-ui, sans-serif', borderLeft: '1px solid #e0e0e0', borderBottom: '1px solid #e0e0e0', boxSizing: 'border-box' },
  header: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', minHeight: 44, padding: '0 10px 0 14px', borderBottom: '1px solid #e5e5e5', fontWeight: 600, fontSize: 14 },
  toggle: { appearance: 'none', display: 'grid', placeItems: 'center', width: 28, height: 28, padding: 0, border: 0, borderRadius: 4, background: 'transparent', color: '#424242', cursor: 'pointer' },
  count: { marginLeft: 8, padding: '1px 8px', borderRadius: 10, background: '#f0f0f0', color: '#424242', fontSize: 12, fontWeight: 400 },
  list: { margin: 0, padding: '6px 6px 10px', overflowY: 'auto', listStyle: 'none' },
  row: { display: 'flex', minHeight: 30, padding: 0 },
  button: { display: 'flex', alignItems: 'flex-start', gap: 8, width: '100%', padding: '6px 8px', border: 0, borderRadius: 3, background: 'transparent', color: 'inherit', font: 'inherit', textAlign: 'left', cursor: 'pointer' },
  dot: { flex: '0 0 auto', width: 8, height: 8, marginTop: 5, borderRadius: '50%' },
  message: { display: 'block', overflow: 'hidden', wordBreak: 'break-word' },
  empty: { margin: '12px 14px', color: '#616161', textAlign: 'center' },
};

export function issueMessage(t: TFunction, issue: ValidationIssue, name: string, otherName: string | null): string {
  switch (issue.rule) {
    case 'dangling-connector':
      return issue.endpoint === null
        ? t('issuesPanel.rules.dangling-connector-unconnected', { name })
        : t('issuesPanel.rules.dangling-connector', { name, endpoint: t(`issuesPanel.endpoint.${issue.endpoint}` as 'issuesPanel.endpoint.begin') });
    case 'isolated-shape':
      return t('issuesPanel.rules.isolated-shape', { name });
    case 'overlapping-shapes':
      return t('issuesPanel.rules.overlapping-shapes', { name, other: otherName ?? name });
    case 'connector-crossing':
      return t('issuesPanel.rules.connector-crossing', { name, other: otherName ?? name });
    case 'empty-shape-data':
      return t('issuesPanel.rules.empty-shape-data', { name, row: issue.row ?? '' });
    default:
      return issue.id;
  }
}

export function IssuesPanel({ issues, collapsed, onToggleCollapsed, onSelect, getShapeName, t, className }: IssuesPanelProps) {
  return (
    <aside className={className} style={styles.root} aria-label={t('issuesPanel.title')}>
      <header style={styles.header}>
        <span>{t('issuesPanel.title')}<span style={styles.count}>{issues.length}</span></span>
        <button
          type="button"
          aria-label={collapsed ? t('issuesPanel.expand') : t('issuesPanel.collapse')}
          aria-expanded={!collapsed}
          title={collapsed ? t('issuesPanel.expand') : t('issuesPanel.collapse')}
          onClick={onToggleCollapsed}
          style={styles.toggle}
        >
          <svg aria-hidden="true" width="16" height="16" viewBox="0 0 16 16">
            <path d={collapsed ? 'M 6 3 L 11 8 L 6 13' : 'M 10 3 L 5 8 L 10 13'} fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </button>
      </header>
      {!collapsed && (issues.length === 0 ? <p style={styles.empty}>{t('issuesPanel.empty')}</p> : (
        <ul style={styles.list}>
          {issues.map((issue) => {
            const name = getShapeName(issue.pageId, issue.shapeId) ?? t('issuesPanel.unnamedShape', { id: issue.shapeId });
            const otherName = issue.otherShapeId ? getShapeName(issue.pageId, issue.otherShapeId) ?? t('issuesPanel.unnamedShape', { id: issue.otherShapeId }) : null;
            return (
              <li key={issue.id} style={styles.row}>
                <button
                  type="button"
                  style={styles.button}
                  aria-label={t('issuesPanel.selectIssue', { name })}
                  onClick={() => onSelect(issue)}
                >
                  <span aria-hidden="true" style={{ ...styles.dot, background: issue.severity === 'error' ? '#c4314b' : '#e6a23c' }} />
                  <span style={styles.message}>{issueMessage(t, issue, name, otherName)}</span>
                </button>
              </li>
            );
          })}
        </ul>
      ))}
    </aside>
  );
}
