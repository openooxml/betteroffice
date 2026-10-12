import { Component, forwardRef, lazy, Suspense, useState, type ReactNode } from 'react';
import type { DocxEditorProps, DocxEditorRef } from '../../DocxEditor';
import { DefaultLoadingIndicator, ParseError } from '../../DocxEditorHelpers';
import { LocaleProvider } from '../../../i18n';

export const docxPeerlessEditorLoader = { load: () => import('./PeerlessDocxEditor') };

export class DocxWorkerEditorChunkLoadError extends Error {
  readonly code = 'chunk-load-failed';
  constructor(readonly cause: unknown) {
    super('Failed to load the worker editor. Remount the editor to retry.');
    this.name = 'DocxWorkerEditorChunkLoadError';
  }
}

type BoundaryProps = Pick<DocxEditorProps, 'onError'> & { children: ReactNode };

class PeerlessEditorBoundary extends Component<BoundaryProps, { error: Error | null }> {
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
    return this.state.error ? <ParseError message={this.state.error.message} /> : this.props.children;
  }
}

export const LazyPeerlessDocxEditor = forwardRef<DocxEditorRef, DocxEditorProps>(function LazyPeerlessDocxEditor(props, ref) {
  const [Editor] = useState(() => lazy(() => docxPeerlessEditorLoader.load()
    .then((module) => ({ default: module.PeerlessDocxEditor }))
    .catch((cause: unknown) => { throw new DocxWorkerEditorChunkLoadError(cause); })));
  return <LocaleProvider i18n={props.i18n}>
    <PeerlessEditorBoundary onError={props.onError}>
      <Suspense fallback={props.loadingIndicator ?? <DefaultLoadingIndicator />}>
        <Editor {...props} ref={ref} />
      </Suspense>
    </PeerlessEditorBoundary>
  </LocaleProvider>;
});
