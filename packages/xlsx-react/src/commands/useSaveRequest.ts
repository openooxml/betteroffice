import { useCallback, useRef } from 'react';
import type { XlsxEditorProps } from '../XlsxEditor';
import { XlsxCommandAdmissionError } from './createXlsxCommandStore';
import type { XlsxCommandResult, XlsxPluginCommandResult } from './types';

export function useSaveRequest(options: {
  document(): object | null;
  onSaveRequest?: XlsxEditorProps['onSaveRequest'];
  fail(error: unknown): void;
}) {
  const latest = useRef(options);
  latest.current = options;
  const pending = useRef<{ document: object; promise: Promise<XlsxPluginCommandResult> } | null>(null);
  return useCallback((save: () => Promise<XlsxPluginCommandResult>): Promise<XlsxPluginCommandResult> => {
    const current = latest.current;
    if (!current.onSaveRequest) return save();
    const document = current.document();
    const replaced = (): XlsxCommandResult => ({
      ok: false, failure: { code: 'document-replaced', message: 'The workbook was replaced' },
    });
    if (!document) return Promise.resolve(replaced());
    if (pending.current?.document === document) return pending.current.promise;
    const promise = Promise.resolve().then(async (): Promise<XlsxPluginCommandResult> => {
      if (latest.current.document() !== document) return replaced();
      const decision = current.onSaveRequest ? await current.onSaveRequest() : true;
      if (latest.current.document() !== document) return replaced();
      return decision === true ? save() : { ok: true, status: 'noop' };
    }).catch((error: unknown): XlsxCommandResult => {
      if (latest.current.document() !== document) return replaced();
      if (error instanceof XlsxCommandAdmissionError) throw error;
      latest.current.fail(error);
      return { ok: false, failure: { code: 'command-failed', message: error instanceof Error ? error.message : String(error) } };
    }).finally(() => {
      if (pending.current?.promise === promise) pending.current = null;
    });
    pending.current = { document, promise };
    return promise;
  }, []);
}
