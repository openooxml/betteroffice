import { Component } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { LocaleProvider } from '../i18n';
import { WorkbookFailure } from '../components/WorkbookFailure';
import { asError } from './asError';
import type { PeerlessSessionWorkbookProps } from './usePeerlessSessionWorkbook';
import { PeerlessXlsxEditorLoader } from './PeerlessXlsxEditorLoader';
import { XlsxWorkerEditorChunkLoadError } from './XlsxWorkerEditorChunkLoadError';
export { XlsxWorkerEditorChunkLoadError } from './XlsxWorkerEditorChunkLoadError';

export const xlsxPeerlessEditorLoader = { load: () => import('./XlsxPeerlessWorkerEditor') };
let prefetched = false;
export function prefetchPeerlessXlsxEditor(): void {
  if (prefetched) return;
  prefetched = true;
  import('./XlsxPeerlessWorkerEditor').catch(() => { prefetched = false; });
}
const reloadPage = () => globalThis.location?.reload();

type PeerlessEditorBoundaryProps = Pick<PeerlessSessionWorkbookProps, 'onError' | 'className' | 'style'> & {
  children: ReactNode;
};

function PeerlessEditorError({ error, className, style }: { error: Error; className?: string; style?: CSSProperties }) {
  const chunk = error instanceof XlsxWorkerEditorChunkLoadError;
  return <WorkbookFailure kind={chunk ? 'load' : 'open'} details={error.message} onReload={chunk ? reloadPage : undefined}
    className={className} style={{ position: 'relative', width: '100%', height: '100%', ...style }} />;
}

class PeerlessEditorBoundary extends Component<PeerlessEditorBoundaryProps, { error: Error | null }> {
  state = { error: null as Error | null };
  private reported: Error | null = null;

  static getDerivedStateFromError(error: unknown) { return { error: asError(error) }; }

  componentDidCatch() {
    const error = this.state.error;
    if (!error || this.reported === error) return;
    this.reported = error;
    try { this.props.onError?.(error); } catch {}
  }

  render() {
    return this.state.error ? <PeerlessEditorError error={this.state.error} className={this.props.className}
      style={this.props.style} /> : this.props.children;
  }
}

type XlsxWorkerEditorProps = PeerlessSessionWorkbookProps & { backend: 'peerless' };

export function XlsxWorkerEditor(props: XlsxWorkerEditorProps) {
  return <LocaleProvider i18n={props.i18n}>
    <PeerlessEditorBoundary onError={props.onError} className={props.className} style={props.style}>
      <PeerlessXlsxEditorLoader load={xlsxPeerlessEditorLoader.load} ErrorType={XlsxWorkerEditorChunkLoadError}
        editorProps={props} />
    </PeerlessEditorBoundary>
  </LocaleProvider>;
}
