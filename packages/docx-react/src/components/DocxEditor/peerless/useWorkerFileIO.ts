import { useCallback, useRef } from 'react';
import type { DocxInput } from '@betteroffice/docx/utils';
import type { DocxEditorProps, DocxEditorRef } from '../../DocxEditor';
import type { DocxCommandId, DocxCommandResult } from '../../../commands/types';

export function useWorkerFileIO(props: DocxEditorProps, ref: DocxEditorRef,
  load: (source: DocxInput) => Promise<void>, identity: () => object | null) {
  const input = useRef<HTMLInputElement>(null);
  const pending = useRef<Promise<DocxCommandResult> | null>(null);
  const latest = useRef({ props, ref, load, identity });
  latest.current = { props, ref, load, identity };
  const perform = useCallback(async (id: DocxCommandId): Promise<DocxCommandResult> => {
    if (id === 'open') { input.current?.click(); return { ok: true, status: 'executed' }; }
    if (id !== 'save') return { ok: false, failure: { code: 'unsupported-command', message: 'Worker command is unavailable.' } };
    if (pending.current) return pending.current;
    const current = latest.current;
    const document = current.identity();
    const job = Promise.resolve().then(async (): Promise<DocxCommandResult> => {
      if (current.props.onSaveRequest && await current.props.onSaveRequest() !== true) return { ok: true, status: 'executed' };
      if (document !== latest.current.identity()) throw new Error('Document was replaced during the save request');
      const bytes = await current.ref.save();
      if (!bytes) throw new Error('Worker save returned no document');
      if (current.props.downloadOnSave !== false) {
        const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
        try {
          const anchor = window.document.createElement('a');
          anchor.href = url;
          anchor.download = `${(current.props.documentName?.trim() || 'document').replace(/\.docx$/i, '')}.docx`;
          anchor.click();
        } finally { setTimeout(() => URL.revokeObjectURL(url), 0); }
      }
      return { ok: true, status: 'executed' };
    }).catch((cause): DocxCommandResult => {
      const error = cause instanceof Error ? cause : new Error(String(cause));
      latest.current.props.onError?.(error);
      return { ok: false, failure: { code: 'command-failed', message: error.message } };
    }).finally(() => { if (pending.current === job) pending.current = null; });
    pending.current = job;
    return job;
  }, []);
  const change = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0];
    event.currentTarget.value = '';
    if (!file) return;
    try {
      const current = latest.current;
      if (current.props.onOpen) await current.props.onOpen(file);
      else { await current.load(file); current.props.onDocumentNameChange?.(file.name.replace(/\.docx$/i, '')); }
    } catch (cause) { latest.current.props.onError?.(cause instanceof Error ? cause : new Error(String(cause))); }
  }, []);
  return { input, change, perform };
}
