import { useEffect, useMemo, useState, type JSX } from 'react';
import type { PptxAsyncEditorProps, PptxEditorCollaborationOptions, PptxEditorProps, PptxInThreadEditorProps, PptxWorkerViewerProps } from './PptxEditor';
import { PptxEditorContent } from './PptxEditorContent';
import { useStableFonts } from './fontFaces';
import { warnPptxEngine, type PptxEngineWarning } from './engineWarnings';
import { PptxSessionViewer } from './viewer/PptxSessionViewer';
import { PptxWorkerEditor, type PptxWorkerEditorProps } from './worker/PptxWorkerEditor';

type WorkerEditorSessionProps = Omit<PptxWorkerEditorProps, 'readOnly'> & { readOnly?: boolean };

function choosePptxEngine(workerOpen: boolean | undefined, collaboration: PptxEditorCollaborationOptions | undefined,
  workerAvailable: boolean): PptxEngineWarning | 'in-thread' | undefined {
  if (workerOpen === false) return 'deprecated';
  if (workerOpen !== true) return 'in-thread';
  if (collaboration != null) return 'collaboration';
  if (!workerAvailable) return 'no-worker';
}

export function PptxEditorSession(props: Exclude<PptxEditorProps, PptxAsyncEditorProps>) {
  useEffect(() => { warnPptxEngine('legacy-sync-api'); }, []);
  const fonts = useStableFonts(props.fonts);
  const workerOpen = props.experimentalWorkerOpen;
  const identity = useMemo(() => ({}), [props.file, fonts, props.clientId, props.collaboration, workerOpen]);
  const engineIdentity = useMemo(() => ({}), [props.file, props.collaboration, workerOpen]);
  const startEngine = () => ({ identity: engineIdentity,
    reason: choosePptxEngine(workerOpen, props.collaboration, typeof Worker === 'function') });
  const [engineChoice, setEngineChoice] = useState(startEngine);
  const engine = engineChoice.identity === engineIdentity ? engineChoice : startEngine();
  if (engineChoice !== engine) setEngineChoice(engine);
  const needsEditor = !props.readOnly || (props.plugins?.length ?? 0) > 0 || props.onChange !== undefined;
  const start = () => ({ identity, editable: needsEditor });
  const [choice, setChoice] = useState(start);
  const session = choice.identity === identity ? choice : start();
  const editable = session.editable || needsEditor;
  if (choice !== session || choice.editable !== editable) setChoice({ ...session, editable });
  useEffect(() => { if (engine.reason && engine.reason !== 'in-thread') warnPptxEngine(engine.reason); }, [engine.reason]);
  const Editor = PptxWorkerEditor as (props: WorkerEditorSessionProps) => JSX.Element;
  return engine.reason ? (
    <PptxEditorContent {...props as PptxInThreadEditorProps} />
  ) : editable ? (
    <Editor {...props as WorkerEditorSessionProps} />
  ) : (
    <PptxSessionViewer {...props as PptxWorkerViewerProps} />
  );
}

