/**
 * DocxEditor Helper Components
 *
 * Small presentational components used by DocxEditor for
 * loading, placeholder, and error states.
 */

import React from 'react';
import { useTranslation } from '../i18n';

// ============================================================================
// HELPER COMPONENTS
// ============================================================================

/**
 * Default loading indicator
 */
export function DefaultLoadingIndicator(): React.ReactElement {
  const { t } = useTranslation();
  return (
    <div
      role="status"
      style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        height: '100%',
        gap: '20px',
        color: 'var(--doc-text-muted)',
      }}
    >
      <div
        style={{
          width: '36px',
          height: '36px',
          border: '3px solid var(--doc-border)',
          borderTopColor: 'var(--doc-primary)',
          borderRadius: '50%',
          animation: 'docx-spin 0.8s linear infinite',
        }}
      />
      <style>
        {`
          @keyframes docx-spin {
            to { transform: rotate(360deg); }
          }
        `}
      </style>
      <div style={{ fontSize: '14px' }}>{t('errors.loadingDocument')}</div>
    </div>
  );
}

/**
 * Default placeholder
 */
export function DefaultPlaceholder(): React.ReactElement {
  const { t } = useTranslation();
  return (
    <div
      style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        height: '100%',
        color: 'var(--doc-text-placeholder)',
      }}
    >
      <svg
        width="64"
        height="64"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.5"
      >
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
        <polyline points="14 2 14 8 20 8" />
      </svg>
      <div style={{ marginTop: '16px' }}>{t('errors.noDocumentLoaded')}</div>
    </div>
  );
}

/**
 * Failure state shown in place of the document: a title, a short
 * explanation, the technical message behind a details disclosure, and an
 * optional action.
 */
function FailureState({ title, description, details, action }: {
  title: string; description: string; details?: string; action?: React.ReactNode;
}): React.ReactElement {
  const { t } = useTranslation();
  return (
    <div
      role="alert"
      style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        height: '100%',
        padding: '20px',
        textAlign: 'center',
      }}
    >
      <div style={{ color: 'var(--doc-error)', marginBottom: '16px' }}>
        <svg
          width="48"
          height="48"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          aria-hidden="true"
        >
          <circle cx="12" cy="12" r="10" />
          <path d="M12 8v4M12 16v.01" />
        </svg>
      </div>
      <h3 style={{ color: 'var(--doc-text)', margin: '0 0 8px' }}>{title}</h3>
      <p style={{ color: 'var(--doc-text-muted)', maxWidth: '400px', margin: '0 0 12px' }}>{description}</p>
      {action}
      {details ? (
        <details style={{ color: 'var(--doc-text-muted)', maxWidth: '480px', marginTop: '12px', fontSize: '12px' }}>
          <summary style={{ cursor: 'pointer' }}>{t('errors.showDetails')}</summary>
          <pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', textAlign: 'start', margin: '8px 0 0' }}>{details}</pre>
        </details>
      ) : null}
    </div>
  );
}

/**
 * Shown when a document could not be opened. The technical message stays
 * behind the details disclosure.
 */
export function ParseError({ message }: { message: string }): React.ReactElement {
  const { t } = useTranslation();
  return <FailureState title={t('errors.failedToLoad')} description={t('errors.failedToLoadDescription')} details={message} />;
}

/**
 * Shown when the editor itself could not load (for example a lazy chunk
 * failed to download). `onRetry` remounts the editor.
 */
export function EditorLoadError({ message, onRetry }: { message: string; onRetry: () => void }): React.ReactElement {
  const { t } = useTranslation();
  return (
    <FailureState
      title={t('errors.editorUnavailable')}
      description={t('errors.editorUnavailableDescription')}
      details={message}
      action={
        <button
          type="button"
          onClick={onRetry}
          style={{
            padding: '6px 16px',
            borderRadius: '6px',
            border: '1px solid var(--doc-border)',
            background: 'var(--doc-surface)',
            color: 'var(--doc-text)',
            cursor: 'pointer',
          }}
        >
          {t('common.retry')}
        </button>
      }
    />
  );
}
