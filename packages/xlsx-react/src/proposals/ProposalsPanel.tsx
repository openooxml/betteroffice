/**
 * The agent-proposals review panel: a labeled popover listing each pending
 * proposal with its agent, note, and per-cell `old → new` previews, plus
 * accept/reject controls. Stale accepts surface an inline warning with a
 * force-apply. Framework chrome only — all state lives in {@link XlsxEditor}.
 */

import type { Proposal } from '@betteroffice/xlsx';
import { en } from '../i18n';
import { proposalColor } from './palette';

/**
 * Props for {@link ProposalsPanel}.
 */
export interface ProposalsPanelProps {
  proposals: Proposal[];
  /** a1 lists keyed by proposal id: the cells that drifted since it was staged. */
  staleFor: Record<string, string[]>;
  onAccept: (id: string, force?: boolean) => void;
  onReject: (id: string) => void;
}

const t = en.proposals;

function fill(template: string, vars: Record<string, string | number>): string {
  return template.replace(/\{(\w+)\}/g, (_, key) => String(vars[key] ?? ''));
}

function cellCountLabel(count: number): string {
  return count === 1 ? t.cellCountOne : fill(t.cellCount, { count });
}

/**
 * The list of pending proposals with review controls.
 */
export function ProposalsPanel({ proposals, staleFor, onAccept, onReject }: ProposalsPanelProps) {
  return (
    <div
      data-testid="xlsx-proposals-panel"
      role="region"
      aria-label={t.panelLabel}
      style={{
        position: 'absolute',
        top: '100%',
        right: 0,
        marginTop: 4,
        width: 320,
        maxHeight: 420,
        overflowY: 'auto',
        background: '#ffffff',
        border: '1px solid #d0d0d0',
        borderRadius: 6,
        boxShadow: '0 6px 24px rgba(0, 0, 0, 0.16)',
        padding: 8,
        zIndex: 10,
        font: '13px system-ui, sans-serif',
        textAlign: 'left',
      }}
    >
      {proposals.length === 0 ? (
        <div style={{ padding: 12, color: '#707070', textAlign: 'center' }}>{t.empty}</div>
      ) : (
        proposals.map((proposal) => {
          const color = proposalColor(proposal.agentId);
          const stale = staleFor[proposal.id];
          return (
            <div
              key={proposal.id}
              data-testid="xlsx-proposal"
              data-proposal-id={proposal.id}
              style={{
                borderLeft: `3px solid ${color}`,
                padding: '8px 10px',
                marginBottom: 8,
                background: '#fafafa',
                borderRadius: 4,
              }}
            >
              <div
                style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}
              >
                <strong style={{ color }}>{proposal.agentId}</strong>
                <span style={{ color: '#707070', fontSize: 12 }}>
                  {cellCountLabel(proposal.cells.length)}
                </span>
              </div>
              {proposal.note && (
                <div style={{ color: '#404040', margin: '4px 0' }}>{proposal.note}</div>
              )}
              <ul style={{ listStyle: 'none', margin: '6px 0', padding: 0 }}>
                {proposal.cells.map((cell) => (
                  <li
                    key={cell.a1}
                    data-testid="xlsx-proposal-cell"
                    style={{ fontVariantNumeric: 'tabular-nums', lineHeight: 1.5 }}
                  >
                    <strong data-testid="xlsx-proposal-cell-a1">{cell.a1}</strong>:{' '}
                    <span style={{ color: '#909090' }}>{cell.oldText || '∅'}</span> {t.changeArrow}{' '}
                    <span data-testid="xlsx-proposal-cell-new" style={{ color }}>
                      {cell.newText}
                    </span>
                  </li>
                ))}
              </ul>
              {stale && stale.length > 0 && (
                <div
                  data-testid="xlsx-proposal-stale"
                  role="alert"
                  style={{
                    color: '#b45309',
                    background: '#fef3c7',
                    borderRadius: 4,
                    padding: '4px 6px',
                    margin: '4px 0',
                    fontSize: 12,
                  }}
                >
                  {fill(t.staleWarning, { cells: stale.join(', ') })}
                </div>
              )}
              <div style={{ display: 'flex', gap: 6, marginTop: 4 }}>
                <button
                  data-testid="xlsx-proposal-accept"
                  onClick={() => onAccept(proposal.id)}
                  style={{ padding: '3px 10px', cursor: 'pointer' }}
                >
                  {t.accept}
                </button>
                {stale && stale.length > 0 && (
                  <button
                    data-testid="xlsx-proposal-force"
                    onClick={() => onAccept(proposal.id, true)}
                    style={{ padding: '3px 10px', cursor: 'pointer', color: '#b45309' }}
                  >
                    {t.forceApply}
                  </button>
                )}
                <button
                  data-testid="xlsx-proposal-reject"
                  onClick={() => onReject(proposal.id)}
                  style={{ padding: '3px 10px', cursor: 'pointer' }}
                >
                  {t.reject}
                </button>
              </div>
            </div>
          );
        })
      )}
    </div>
  );
}
