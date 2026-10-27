import type { CSSProperties, RefObject } from 'react';
import { toolbarFocusStyle, toolbarTheme, useFocusRing } from '../../../../shared/react-toolbar/theme';
import { usePptxCommands, usePptxCommandState } from '../commands/hooks';
import { useTranslation } from '../i18n';
import { useDisabledDescription } from './ui/ToolbarPrimitives';

export function ProposalsButton({
  buttonRef,
  count,
}: {
  buttonRef?: RefObject<HTMLButtonElement | null>;
  count: number;
}) {
  const store = usePptxCommands();
  const state = usePptxCommandState('proposalsPanel');
  const { t } = useTranslation();
  const described = useDisabledDescription(
    !state.enabled,
    state.enabled ? undefined : state.disabledReason?.message
  );
  const focus = useFocusRing<HTMLButtonElement>();
  return (
    <>
      <button
        ref={buttonRef}
        type="button"
        data-testid="pptx-proposals-button"
        {...described.props}
        aria-expanded={state.value ?? false}
        title={described.reason}
        {...focus.focusProps}
        style={{ ...proposalsButton, ...toolbarFocusStyle(focus.focusRing) }}
        onClick={() => {
          if (state.enabled) void store.execute('proposalsPanel', null);
        }}
      >
        {t('proposals.title')} <span data-testid="pptx-proposals-count">{count}</span>
      </button>
      {described.node}
    </>
  );
}

const proposalsButton: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  gap: 6,
  flex: '0 0 auto',
  margin: '0 0 0 4px',
  padding: '0 10px',
  height: 28,
  border: `1px solid ${toolbarTheme.line}`,
  borderRadius: 6,
  background: toolbarTheme.surface,
  color: toolbarTheme.ink,
  font: '500 13px ui-sans-serif, system-ui, sans-serif',
  cursor: 'pointer',
  boxSizing: 'border-box',
};
