import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { PptxEditorContent, type PptxAsyncEditorInternalProps, type PptxEditorApi } from './PptxEditor';
import type { PptxAsyncEditorApi } from './worker/createPeerlessEditorApi';
import { createInThreadAsyncEditorApi, asyncCollaboration } from './createInThreadAsyncEditorApi';
import { generateCollaborationClientId } from './collaborationClientId';

export function PptxAsyncInThreadSession(props: PptxAsyncEditorInternalProps) {
  const callbacks = useRef(props);
  callbacks.current = props;
  const [api, setApi] = useState<PptxAsyncEditorApi | null>(null);
  const readOnly = useRef(props.readOnly);
  readOnly.current = props.readOnly;
  const ready = useCallback((value: PptxEditorApi) => {
    setApi(createInThreadAsyncEditorApi(value, () => Boolean(readOnly.current)));
  }, []);
  const collaborationIdentity = useMemo(() => ({}), [Boolean(props.collaboration), props.collaboration?.roomId]);
  const generatedClientId = useMemo(() => props.collaboration ? generateCollaborationClientId() : 0, [collaborationIdentity]);
  const onReplica = useMemo(() => asyncCollaboration(props.collaboration, generatedClientId)?.onReplica,
    [props.collaboration?.onReplica]);
  const collaboration = props.collaboration
    ? { ...props.collaboration, clientId: props.collaboration.clientId ?? generatedClientId, onReplica } : undefined;
  useEffect(() => {
    if (!api) return;
    const cleanup = callbacks.current.onReady?.(api);
    return typeof cleanup === 'function' ? cleanup : undefined;
  }, [api]);
  return <PptxEditorContent {...props} collaboration={collaboration} collaborationIdentity={collaborationIdentity} onReady={ready} />;
}
