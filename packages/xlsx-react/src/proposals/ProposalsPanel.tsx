/**
 * The agent-proposals review panel: a labeled popover listing each pending
 * proposal with its agent, note, and per-cell `old → new` previews, plus
 * accept/reject controls bound to the editor's proposal commands. Stale
 * accepts surface an inline warning with a force-apply.
 */

import type { Proposal } from '@betteroffice/xlsx';
import type { CSSProperties } from 'react';
import { xlsxCommandController } from '../commands/createXlsxCommandStore';
import { useXlsxCommands, useXlsxCommandState } from '../commands/hooks';
import type { XlsxCommandArgs, XlsxCommandResult, XlsxCommandStore } from '../commands/types';
import { useTranslation } from '../i18n';
import { proposalColor } from './palette';

/**
 * Props for {@link ProposalsPanel}.
 */
export interface ProposalsPanelProps {
  proposals: Proposal[];
  /** a1 lists keyed by proposal id: the cells that drifted since it was staged. */
  staleFor: Record<string, string[]>;
  style?: CSSProperties;
}

function run<K extends 'proposalAccept' | 'proposalReject'>(
  store: XlsxCommandStore,
  id: K,
  args: XlsxCommandArgs[K]
): void {
  void store.execute(id, args).then((result: XlsxCommandResult) => {
    const focus = document.activeElement;
    if (result.ok && (!focus || focus === document.body)) {
      xlsxCommandController(store)?.focusEditor();
    }
  });
}

function ReviewButton({
  id,
  args,
  label,
  testId,
  color,
}: {
  id: 'proposalAccept' | 'proposalReject';
  args: XlsxCommandArgs['proposalAccept'];
  label: string;
  testId: string;
  color?: string;
}) {
  const store = useXlsxCommands();
  const state = useXlsxCommandState(id, args);
  const reason = state.enabled ? undefined : state.disabledReason.message;
  return (
    <button
      data-testid={testId}
      aria-disabled={reason ? true : undefined}
      title={reason}
      onClick={() => {
        if (state.enabled) run(store, id, args);
      }}
      style={{
        padding: '3px 10px',
        cursor: state.enabled ? 'pointer' : 'default',
        color,
        opacity: state.enabled ? 1 : 0.48,
      }}
    >
      {label}
    </button>
  );
}

/**
 * The list of pending proposals with review controls.
 */
export function ProposalsPanel({ proposals, staleFor, style }: ProposalsPanelProps) {
  const { t } = useTranslation();
  return (
    <div
      data-testid="xlsx-proposals-panel"
      role="region"
      aria-label={t('proposals.panelLabel')}
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
        borderRadius: 10,
        boxShadow: '0 6px 24px rgba(0, 0, 0, 0.16)',
        padding: 8,
        zIndex: 10,
        font: '13px system-ui, sans-serif',
        textAlign: 'left',
        ...style,
      }}
    >
      {proposals.length === 0 ? (
        <div style={{ padding: 12, color: '#707070', textAlign: 'center' }}>
          {t('proposals.empty')}
        </div>
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
                  {t('proposals.cellCount', { count: proposal.cells.length })}
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
                    <span style={{ color: '#909090' }}>{cell.oldText || '∅'}</span>{' '}
                    {t('proposals.changeArrow')}{' '}
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
                  {t('proposals.staleWarning', { cells: stale.join(', ') })}
                </div>
              )}
              <div style={{ display: 'flex', gap: 6, marginTop: 4 }}>
                <ReviewButton
                  id="proposalAccept"
                  args={{ proposalId: proposal.id }}
                  label={t('proposals.accept')}
                  testId="xlsx-proposal-accept"
                />
                {stale && stale.length > 0 && (
                  <ReviewButton
                    id="proposalAccept"
                    args={{ proposalId: proposal.id, force: true }}
                    label={t('proposals.forceApply')}
                    testId="xlsx-proposal-force"
                    color="#b45309"
                  />
                )}
                <ReviewButton
                  id="proposalReject"
                  args={{ proposalId: proposal.id }}
                  label={t('proposals.reject')}
                  testId="xlsx-proposal-reject"
                />
              </div>
            </div>
          );
        })
      )}
    </div>
  );
}
