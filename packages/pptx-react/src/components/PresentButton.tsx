import type { CSSProperties } from 'react';
import {
  TOOLBAR_DISABLED_OPACITY,
  toolbarFocusStyle,
  toolbarTheme,
  useFocusRing,
} from '../../../../shared/react-toolbar/theme';
import { usePptxCommands, usePptxCommandState } from '../commands/hooks';
import { useTranslation } from '../i18n';
import { useDisabledDescription } from './ui/ToolbarPrimitives';

export function PresentButton() {
  const store = usePptxCommands();
  const state = usePptxCommandState('slideshow');
  const { t } = useTranslation();
  const described = useDisabledDescription(
    !state.enabled,
    state.enabled ? undefined : state.disabledReason?.message
  );
  const focus = useFocusRing<HTMLButtonElement>();
  return (
    <>
      <button
        type="button"
        onClick={() => {
          if (state.enabled) void store.execute('slideshow', null);
        }}
        {...described.props}
        title={described.reason}
        data-testid="pptx-present"
        {...focus.focusProps}
        style={{
          ...presentButton,
          opacity: state.enabled ? 1 : TOOLBAR_DISABLED_OPACITY,
          ...toolbarFocusStyle(focus.focusRing),
        }}
      >
        <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
          <path d="M6 4.5v15l13-7.5Z" />
        </svg>
        {t('toolbar.present')}
      </button>
      {described.node}
    </>
  );
}

const presentButton: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  gap: 6,
  flex: '0 0 auto',
  margin: '0 12px 0 4px',
  padding: '0 12px',
  height: 28,
  border: 0,
  borderRadius: 6,
  background: toolbarTheme.accent,
  color: toolbarTheme.accentInk,
  font: '500 13px ui-sans-serif, system-ui, sans-serif',
  cursor: 'pointer',
};
