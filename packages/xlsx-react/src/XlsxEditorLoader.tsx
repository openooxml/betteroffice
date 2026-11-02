import { Component, lazy, Suspense, useState } from 'react';
import type { ReactNode } from 'react';
import type { XlsxAsyncEditorProps } from './XlsxEditor.types';
import type { LegacyXlsxEditor } from './LegacyXlsxEditor';
import type { XlsxAsyncInThreadSession } from './XlsxAsyncInThreadSession';
import { WorkbookFailure } from './components/WorkbookFailure';
import { XlsxWorkerEditorChunkLoadError } from './worker/XlsxWorkerEditorChunkLoadError';

export const xlsxEditorLoaders = {
  legacy: () => import('./LegacyXlsxEditor').then((module) => ({ default: module.LegacyXlsxEditor })),
  inThread: () => import('./XlsxAsyncInThreadSession').then((module) => ({ default: module.XlsxAsyncInThreadSession })),
};

type BoundaryProps = Pick<XlsxAsyncEditorProps, 'file' | 'loadingIndicator' | 'onError' | 'className' | 'style'> & {
  children: ReactNode;
};
const reloadPage = () => globalThis.location?.reload();

class XlsxEditorBoundary extends Component<BoundaryProps, { error: Error | null }> {
  state = { error: null as Error | null };
  private reported: Error | null = null;

  static getDerivedStateFromError(error: unknown) {
    return { error: error instanceof Error ? error : new Error(String(error)) };
  }

  componentDidCatch() {
    const error = this.state.error;
    if (!error || this.reported === error) return;
    this.reported = error;
    try { this.props.onError?.(error); } catch {}
  }

  render() {
    const { error } = this.state;
    const { file, loadingIndicator, className, style, children } = this.props;
    const chunk = error instanceof XlsxWorkerEditorChunkLoadError;
    return error ? <WorkbookFailure kind={chunk ? 'load' : 'open'} details={error.message}
      onReload={chunk ? reloadPage : undefined} className={className}
      style={{ position: 'relative', width: '100%', height: '100%', ...style }} />
      : <Suspense fallback={file ? loadingIndicator ?? null : null}>{children}</Suspense>;
  }
}

function chunkFailed(cause: unknown): never { throw new XlsxWorkerEditorChunkLoadError(cause); }

export function LegacyXlsxEditorLoader(props: Parameters<typeof LegacyXlsxEditor>[0]) {
  const [Editor] = useState(() => lazy(() => Promise.resolve().then(xlsxEditorLoaders.legacy).catch(chunkFailed)));
  return <XlsxEditorBoundary {...props}><Editor {...props} /></XlsxEditorBoundary>;
}

export function XlsxAsyncInThreadSessionLoader(props: Parameters<typeof XlsxAsyncInThreadSession>[0]) {
  const [Editor] = useState(() => lazy(() => Promise.resolve().then(xlsxEditorLoaders.inThread).catch(chunkFailed)));
  return <XlsxEditorBoundary {...props}><Editor {...props} /></XlsxEditorBoundary>;
}
