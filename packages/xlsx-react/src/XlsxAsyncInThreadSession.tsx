import { useCallback, useRef } from 'react';
import type { XlsxAsyncEditorProps, XlsxEditorApi } from './XlsxEditor.types';
import { XlsxEditorContent } from './XlsxEditorContent';
import { XlsxCommandAdmissionError } from './commands/createXlsxCommandStore';
import { createInThreadAsyncEditorApi } from './createInThreadAsyncEditorApi';

export function XlsxAsyncInThreadSession(props: Omit<XlsxAsyncEditorProps, 'file'> & { file?: Uint8Array }) {
  const latest = useRef(props);
  latest.current = props;
  const current = useRef<XlsxEditorApi | null>(null);
  const ready = useCallback((api: XlsxEditorApi) => {
    current.current = api;
    const cleanup = props.onReady?.(createInThreadAsyncEditorApi(api, props.collaboration !== undefined, () => {
      if (current.current !== api) throw new XlsxCommandAdmissionError('document-replaced');
    }, () => latest.current.onSave));
    return () => {
      if (current.current === api) current.current = null;
      cleanup?.();
    };
  }, [props.onReady, props.collaboration]);
  return <XlsxEditorContent {...props} experimentalWorkerOpen={false} onReady={ready} />;
}
