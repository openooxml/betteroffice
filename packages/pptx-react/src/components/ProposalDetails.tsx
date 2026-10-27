import type { CSSProperties } from 'react';
import type { ProposalChange } from '@betteroffice/pptx';
import type { TFunction } from '@betteroffice/pptx-i18n';

interface Props { change: ProposalChange; t: TFunction; textOnly?: boolean; testId?: string }

function ProposalTextDiff({ change, t, testId }: Props) {
  if (change.oldText === change.newText) return null;
  return (
    <div
      data-testid={testId}
      style={styles.textDiff}
    >
      <del style={styles.deleted}>{change.oldText || t('proposals.blank')}</del>
      <ins style={styles.inserted}>{change.newText || t('proposals.blank')}</ins>
    </div>
  );
}

export function ProposalChangeDetails({ change, t, textOnly, testId }: Props) {
  return (
    <>
      <ProposalTextDiff change={change} t={t} testId={testId} />
      {!textOnly && change.before && change.after && <ShapeChange change={change} t={t} />}
    </>
  );
}

function ShapeChange({ change, t }: Props) {
  const before = change.before!;
  const after = change.after!;
  const geometryChanged =
    before.x !== after.x ||
    before.y !== after.y ||
    before.width !== after.width ||
    before.height !== after.height;
  const rectangle = (shape: typeof before) =>
    [shape.x, shape.y, shape.width, shape.height]
      .map((value) => (value / 12700).toFixed(1))
      .join(', ');
  return (
    <>
      {geometryChanged && (
        <p style={styles.detail}>
          {t('proposals.geometry')}
          <br />
          {rectangle(before)} → {rectangle(after)} pt
        </p>
      )}
      {JSON.stringify({ ...before, x: 0, y: 0, width: 0, height: 0 }) !==
        JSON.stringify({ ...after, x: 0, y: 0, width: 0, height: 0 }) &&
        change.oldText === change.newText && (
          <p style={styles.detail}>{t('proposals.appearance')}</p>
        )}
    </>
  );
}

const styles: Record<string, CSSProperties> = {
  textDiff: {
    display: 'grid',
    gap: 6,
    margin: '8px 0',
    whiteSpace: 'pre-wrap',
    overflowWrap: 'anywhere',
    maxHeight: 200,
    overflowY: 'auto',
  },
  deleted: { color: '#9f1239', background: '#fff1f2', padding: '4px 6px' },
  inserted: {
    color: '#166534',
    background: '#f0fdf4',
    padding: '4px 6px',
    textDecoration: 'none',
  },
  detail: { color: '#64748b', lineHeight: 1.5 },
};
