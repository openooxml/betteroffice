import { Component, forwardRef, useEffect, useMemo, useState, type ForwardedRef, type ReactNode } from 'react';
import type { DocxEditorProps, DocxEditorRef } from '../../DocxEditor';
import { DefaultLoadingIndicator, ParseError } from '../../DocxEditorHelpers';
import { LocaleProvider } from '../../../i18n';
import type { PeerlessDocxEditor } from './PeerlessDocxEditor';

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

type Load = typeof docxPeerlessEditorLoader.load;
type Request = { load: Load; promise: ReturnType<Load> };
type Chunk = { Editor: typeof PeerlessDocxEditor } | { error: Error } | { loading: true };
const editors = new WeakMap<Load, typeof PeerlessDocxEditor>();
const imports = new WeakMap<Load, ReturnType<Load>>();

function importChunk(load: Load): ReturnType<Load> {
  const inflight = imports.get(load);
  if (inflight) return inflight;
  let started: ReturnType<Load>;
  try { started = load(); }
  catch (cause) { started = Promise.reject(cause); }
  const promise = started;
  imports.set(load, promise);
  const settle = () => { if (imports.get(load) === promise) imports.delete(load); };
  void promise.then((module) => { editors.set(load, module.PeerlessDocxEditor); settle(); }, settle);
  return promise;
}

function usePeerlessChunk(load: Load): Chunk {
  const cached = editors.get(load);
  const request = useMemo(() => editors.has(load) ? null : { load, promise: importChunk(load) }, [load]);
  const [settled, setSettled] = useState<{ request: Request; chunk: Chunk } | null>(null);
  useEffect(() => {
    if (!request) return;
    let active = true;
    void request.promise.then((module) => {
      if (active) setSettled({ request, chunk: { Editor: module.PeerlessDocxEditor } });
    }, (cause: unknown) => {
      if (active) setSettled({ request, chunk: { error: new DocxWorkerEditorChunkLoadError(cause) } });
    });
    return () => { active = false; };
  }, [request]);
  if (cached) return { Editor: cached };
  return settled?.request === request ? settled.chunk : { loading: true };
}

function PeerlessDocxEditorChunk({ editorProps, editorRef }: { editorProps: DocxEditorProps; editorRef: ForwardedRef<DocxEditorRef> }) {
  const chunk = usePeerlessChunk(docxPeerlessEditorLoader.load);
  if ('error' in chunk) throw chunk.error;
  if ('Editor' in chunk) return <chunk.Editor {...editorProps} ref={editorRef} />;
  return <>{editorProps.loadingIndicator ?? <DefaultLoadingIndicator />}</>;
}

export const LazyPeerlessDocxEditor = forwardRef<DocxEditorRef, DocxEditorProps>(function LazyPeerlessDocxEditor(props, ref) {
  return <LocaleProvider i18n={props.i18n}>
    <PeerlessEditorBoundary onError={props.onError}>
      <PeerlessDocxEditorChunk editorProps={props} editorRef={ref} />
    </PeerlessEditorBoundary>
  </LocaleProvider>;
});
