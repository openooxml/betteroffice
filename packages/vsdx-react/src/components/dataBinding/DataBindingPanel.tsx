import type { ShapeSnapshot } from '@betteroffice/vsdx';
import type { TFunction } from '@betteroffice/vsdx-i18n';
import { useEffect, useMemo, useState } from 'react';
import type { CSSProperties } from 'react';
import { matchColumnsToRows, type BindOutcome, type ImportedTable } from './bindTable';

export interface DataBindingPanelProps {
  table: ImportedTable | null;
  shape: ShapeSnapshot | null;
  onImport: () => void;
  onBind: (rowIndex: number) => BindOutcome | void;
  t: TFunction;
  className?: string;
}

const styles: Record<string, CSSProperties> = {
  root: { display: 'flex', flexDirection: 'column', width: 264, minWidth: 264, height: '100%', background: '#fff', color: '#242424', font: '400 13px ui-sans-serif, system-ui, sans-serif', borderLeft: '1px solid #e0e0e0', boxSizing: 'border-box' },
  header: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, minHeight: 44, padding: '0 14px', borderBottom: '1px solid #e5e5e5', fontWeight: 600, fontSize: 14 },
  action: { appearance: 'none', height: 26, padding: '0 10px', border: '1px solid #bdbdbd', borderRadius: 3, background: '#fff', color: 'inherit', font: 'inherit', cursor: 'pointer' },
  note: { margin: '16px 14px', color: '#616161' },
  match: { margin: '10px 14px 0', color: '#616161', fontSize: 12 },
  list: { display: 'flex', flexDirection: 'column', gap: 6, margin: 0, padding: '10px 14px', overflowY: 'auto', listStyle: 'none' },
  row: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, minWidth: 0 },
  cells: { overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis' },
  status: { margin: '0 14px 12px', color: '#616161', fontSize: 12 },
  refusal: { margin: '0 14px 12px', color: '#8b1e2d', fontSize: 12 },
};

export function DataBindingPanel({ table, shape, onImport, onBind, t, className }: DataBindingPanelProps) {
  const [status, setStatus] = useState<{ text: string; refused: boolean } | null>(null);
  const bindings = useMemo(() => (table ? matchColumnsToRows(table, shape) : []), [table, shape]);
  useEffect(() => { setStatus(null); }, [table, shape?.id]);

  const bind = (rowIndex: number) => {
    const outcome = onBind(rowIndex);
    if (!outcome) return;
    if (outcome.refusals.length > 0) {
      const rows = outcome.refusals.map((receipt) => receipt.rowName ?? String(receipt.rowIndex ?? '')).join(', ');
      setStatus({ text: t('dataBinding.refused', { rows }), refused: true });
      return;
    }
    if (!outcome.applied) {
      setStatus({ text: t('dataBinding.nothingToBind'), refused: false });
      return;
    }
    setStatus({ text: t('dataBinding.bound', { count: outcome.receipts.length }), refused: false });
  };

  return (
    <aside className={className} style={styles.root} aria-label={t('dataBinding.title')}>
      <header style={styles.header}>
        <span>{t('dataBinding.title')}</span>
        <button type="button" style={styles.action} onClick={onImport}>{t('dataBinding.import')}</button>
      </header>
      {!table && <p style={styles.note}>{t('dataBinding.noTable')}</p>}
      {table && !shape && <p style={styles.note}>{t('dataBinding.noSelection')}</p>}
      {table && shape && bindings.length === 0 && <p style={styles.note}>{t('dataBinding.unmatched')}</p>}
      {table && shape && bindings.length > 0 && (
        <>
          <p style={styles.match}>{t('dataBinding.matched', { count: bindings.length, total: table.headers.length })}</p>
          <ul style={styles.list}>
            {table.rows.map((row, rowIndex) => (
              <li key={rowIndex} style={styles.row}>
                <span style={styles.cells} title={row.join(' · ')}>{row.join(' · ')}</span>
                <button type="button" style={styles.action} onClick={() => bind(rowIndex)}>{t('dataBinding.bind')}</button>
              </li>
            ))}
          </ul>
        </>
      )}
      {status && <p style={status.refused ? styles.refusal : styles.status} role={status.refused ? 'alert' : 'status'}>{status.text}</p>}
    </aside>
  );
}
