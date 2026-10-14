import { useCallback, useRef } from 'react';
import type { DocxInput } from '@betteroffice/docx/utils';
import type { DocxEditorProps, DocxEditorRef } from '../../DocxEditor';
import type { DocxCommandId, DocxCommandResult } from '../../../commands/types';

export function useWorkerFileIO(props: DocxEditorProps, ref: DocxEditorRef,
  load: (source: DocxInput) => Promise<void>, identity: () => object | null) {
  const input = useRef<HTMLInputElement>(null);
  const imageInput = useRef<HTMLInputElement>(null);
  const imageInsert = useRef<((image: Readonly<Record<string, unknown>>) => Promise<DocxCommandResult>) | null>(null);
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
  const pickImage = useCallback((insert: NonNullable<typeof imageInsert.current>) => {
    imageInsert.current = insert;
    imageInput.current?.click();
  }, []);
  const imageChange = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.currentTarget.files?.[0];
    const insert = imageInsert.current;
    imageInsert.current = null;
    event.currentTarget.value = '';
    if (!file || !insert) return;
    const report = (error: Error) => latest.current.props.onError?.(error);
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result !== 'string') return;
      const image = new Image();
      image.onload = () => {
        let width = image.naturalWidth;
        let height = image.naturalHeight;
        if (width > 612) {
          height = Math.round(height * (612 / width));
          width = 612;
        }
        void insert({ src: reader.result, alt: file.name, width, height, wrapType: 'inline', displayMode: 'inline',
          distTop: 0, distBottom: 0, distLeft: 0, distRight: 0 }).then(result => {
          if (result.ok && result.status === 'executed') latest.current.ref.focus();
        }).catch(cause => report(cause instanceof Error ? cause : new Error(String(cause))));
      };
      image.onerror = () => report(new Error('Failed to decode image'));
      image.src = reader.result;
    };
    reader.onerror = () => report(reader.error ?? new Error('Failed to read image'));
    reader.readAsDataURL(file);
  }, []);
  return { input, change, perform, imageInput, pickImage, imageChange };
}
