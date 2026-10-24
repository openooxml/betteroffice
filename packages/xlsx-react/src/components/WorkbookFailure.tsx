import type { CSSProperties } from 'react';
import { useTranslation } from '../i18n';

/** `open`: the file never opened; `stopped`: the editor failed after opening; `load`: the editor code did not load. */
export type WorkbookFailureKind = 'open' | 'stopped' | 'load';

const COPY = {
  open: ['editor.openError', 'editor.openErrorDescription'],
  stopped: ['editor.stoppedError', 'editor.stoppedErrorDescription'],
  load: ['editor.loadError', 'editor.loadErrorDescription'],
} as const;

const overlay: CSSProperties = {
  position: 'absolute',
  inset: 0,
  zIndex: 3,
  justifyContent: 'center',
  padding: 24,
  background: '#ffffff',
};

const banner: CSSProperties = {
  position: 'relative',
  flex: '0 0 auto',
  padding: '10px 16px',
  borderTop: '1px solid #e0e0e0',
  background: '#fafafa',
};

/**
 * A failure in place of the workbook: a title, a short explanation, an
 * optional Reload page, and the technical message behind a details disclosure.
 */
export function WorkbookFailure({ kind, details, onReload, className, style }: {
  kind: WorkbookFailureKind;
  details?: string;
  onReload?: () => void;
  className?: string;
  style?: CSSProperties;
}) {
  const { t } = useTranslation();
  const [title, description] = COPY[kind];
  return (
    <div
      data-testid="xlsx-error"
      data-failure={kind}
      role="alert"
      className={className}
      style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        gap: 8,
        boxSizing: 'border-box',
        textAlign: 'center',
        color: '#202124',
        font: '14px ui-sans-serif, system-ui, sans-serif',
        ...(kind === 'stopped' ? banner : overlay),
        ...style,
      }}
    >
      <strong style={{ fontSize: kind === 'stopped' ? 14 : 16, fontWeight: 600 }}>{t(title)}</strong>
      <span style={{ maxWidth: 420, color: '#5f6368' }}>{t(description)}</span>
      {onReload && (
        <button
          type="button"
          onClick={onReload}
          style={{
            marginTop: 4,
            padding: '6px 16px',
            border: '1px solid #dadce0',
            borderRadius: 6,
            background: '#ffffff',
            color: '#202124',
            font: 'inherit',
            cursor: 'pointer',
          }}
        >
          {t('editor.reloadPage')}
        </button>
      )}
      {details && (
        <details style={{ maxWidth: 480, color: '#5f6368', fontSize: 12 }}>
          <summary style={{ cursor: 'pointer' }}>{t('editor.showDetails')}</summary>
          <pre style={{ margin: '8px 0 0', whiteSpace: 'pre-wrap', wordBreak: 'break-word', textAlign: 'start' }}>
            {details}
          </pre>
        </details>
      )}
    </div>
  );
}
