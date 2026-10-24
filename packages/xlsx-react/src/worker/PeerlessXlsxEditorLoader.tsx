import { useEffect, useRef, useState } from 'react';
import type { PeerlessSessionWorkbookProps } from './usePeerlessSessionWorkbook';
import type { XlsxPeerlessWorkerEditor } from './XlsxPeerlessWorkerEditor';

type Editor = typeof XlsxPeerlessWorkerEditor;
type Load = () => Promise<{ XlsxPeerlessWorkerEditor: Editor }>;
type Chunk = { Component: Editor } | { error: Error } | { loading: true };
type ErrorType = new (cause: unknown) => Error;
const editors = new WeakMap<Load, Editor>();

function usePeerlessChunk(load: Load, ErrorType: ErrorType): Chunk {
  const cached = editors.get(load);
  const initial: Chunk = cached ? { Component: cached } : { loading: true };
  const [state, setState] = useState<{ load: Load; chunk: Chunk }>(() => ({ load, chunk: initial }));
  const pending = useRef<{ load: Load; promise: ReturnType<Load> } | null>(null);
  useEffect(() => {
    if (cached) return;
    let active = true;
    if (!pending.current || pending.current.load !== load) {
      let promise: ReturnType<Load>;
      try { promise = load(); }
      catch (cause) { promise = Promise.reject(cause); }
      pending.current = { load, promise };
    }
    const request = pending.current;
    void request.promise.then((module) => {
      if (!active || pending.current !== request) return;
      editors.set(load, module.XlsxPeerlessWorkerEditor);
      setState({ load, chunk: { Component: module.XlsxPeerlessWorkerEditor } });
    }).catch((cause: unknown) => {
      if (!active || pending.current !== request) return;
      editors.delete(load);
      setState({ load, chunk: { error: new ErrorType(cause) } });
    });
    return () => { active = false; };
  }, [load, ErrorType]);
  return cached ? { Component: cached } : state.load === load ? state.chunk : { loading: true };
}

export function PeerlessXlsxEditorLoader({ load, ErrorType, editorProps }: {
  load: Load;
  ErrorType: ErrorType;
  editorProps: PeerlessSessionWorkbookProps;
}) {
  const chunk = usePeerlessChunk(load, ErrorType);
  if ('error' in chunk) throw chunk.error;
  return 'Component' in chunk ? <chunk.Component {...editorProps} /> : editorProps.file ? editorProps.loadingIndicator ?? null : null;
}
