import { useEffect, useRef, useState, type ReactNode } from 'react';
import type { PptxAsyncEditorProps } from './PptxEditor';
import type { PeerlessPptxSurface } from './worker/PeerlessPptxSurface';

type Load = () => Promise<{ PeerlessPptxSurface: typeof PeerlessPptxSurface }>;
type Chunk = { Component: typeof PeerlessPptxSurface } | { error: Error } | { loading: true };
type ErrorType = new (cause: unknown) => Error;
const surfaces = new WeakMap<Load, typeof PeerlessPptxSurface>();

function usePeerlessChunk(load: Load, identity: object, ErrorType: ErrorType): Chunk {
  const cached = surfaces.get(load);
  const initial: Chunk = cached ? { Component: cached } : { loading: true };
  const [state, setState] = useState<{ load: Load; identity: object; chunk: Chunk }>(() => ({ load, identity, chunk: initial }));
  const pending = useRef<{ load: Load; identity: object; promise: ReturnType<Load> } | null>(null);
  useEffect(() => {
    if (cached) return;
    let active = true;
    if (!pending.current || pending.current.load !== load || pending.current.identity !== identity) {
      let promise: ReturnType<Load>;
      try { promise = load(); }
      catch (cause) { promise = Promise.reject(cause); }
      pending.current = { load, identity, promise };
    }
    const request = pending.current;
    void request.promise.then((module) => {
      if (!active || pending.current !== request) return;
      surfaces.set(load, module.PeerlessPptxSurface);
      setState({ load, identity, chunk: { Component: module.PeerlessPptxSurface } });
    }).catch((cause: unknown) => {
      if (!active || pending.current !== request) return;
      surfaces.delete(load);
      setState({ load, identity, chunk: { error: new ErrorType(cause) } });
    });
    return () => { active = false; };
  }, [load, identity, ErrorType]);
  return cached ? { Component: cached }
    : state.load === load && state.identity === identity ? state.chunk : { loading: true };
}

export function PeerlessPptxEditorLoader({ load, identity, ErrorType, fallback, editorProps }: {
  load: Load;
  identity: object;
  ErrorType: ErrorType;
  fallback: ReactNode;
  editorProps: PptxAsyncEditorProps;
}) {
  const chunk = usePeerlessChunk(load, identity, ErrorType);
  if ('error' in chunk) throw chunk.error;
  if ('Component' in chunk) return <chunk.Component {...editorProps} />;
  return fallback;
}
