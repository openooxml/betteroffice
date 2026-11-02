import { useEffect, useMemo, useState } from 'react';
import type {
  GeneralEditorSessionProps, WorkerSessionProps, XlsxInThreadEditorProps, XlsxLegacyEditorProps,
  XlsxWorkerEditorProps, XlsxWorkerViewerProps,
} from './XlsxEditor.types';
import { XlsxEditorContent } from './XlsxEditorContent';
import { warnXlsxEngine, type XlsxEngineWarning } from './engineWarnings';
import { XlsxSessionViewer } from './viewer/XlsxSessionViewer';
import { LegacyXlsxWorkerEditor } from './worker/LegacyXlsxWorkerEditor';
import type { EditableSessionWorkbookProps } from './worker/useEditableSessionWorkbook';

export function LegacyXlsxEditor(props: GeneralEditorSessionProps | XlsxLegacyEditorProps | XlsxWorkerViewerProps | XlsxWorkerEditorProps | WorkerSessionProps): React.JSX.Element {
  useEffect(() => { warnXlsxEngine('legacy-sync-api'); }, []);
  const workerOpen = props.experimentalWorkerOpen;
  const collaboration = props.collaboration;
  const identity = useMemo(() => ({}), [props.file, collaboration, workerOpen]);
  const start = () => ({ identity, reason: workerOpen === false ? 'deprecated' as XlsxEngineWarning
    : workerOpen !== true ? 'in-thread' as const
    : typeof Worker !== 'function' ? 'no-worker' as XlsxEngineWarning : undefined, editor: !props.readOnly });
  const [choice, setChoice] = useState(start);
  const session = choice.identity === identity ? choice : start();
  const editor = session.editor || !props.readOnly;
  if (choice !== session || choice.editor !== editor) setChoice({ ...session, editor });
  useEffect(() => { if (session.reason && session.reason !== 'in-thread') warnXlsxEngine(session.reason); }, [session.reason]);
  return session.reason ? <XlsxEditorContent {...props as Omit<XlsxInThreadEditorProps, 'file'> & { file?: Uint8Array }} />
    : editor ? <LegacyXlsxWorkerEditor {...props as EditableSessionWorkbookProps} />
    : <XlsxSessionViewer {...props as XlsxWorkerViewerProps & { file?: Uint8Array }} />;
}
