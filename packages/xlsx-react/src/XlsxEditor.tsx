import { useEffect, useMemo, useState } from 'react';
import type {
  WorkerSessionProps, XlsxAsyncEditorProps, XlsxEditorProps, XlsxLegacyEditorProps,
  XlsxWorkerEditorProps, XlsxWorkerViewerProps,
} from './XlsxEditor.types';
import { LegacyXlsxEditorLoader, XlsxAsyncInThreadSessionLoader } from './XlsxEditorLoader';
import { LocaleProvider } from './i18n';
import { LEGACY_SYNC_API } from './runtimeOptions';
import { warnXlsxEngine } from './engineWarnings';
import { prefetchPeerlessXlsxEditor, XlsxWorkerEditor } from './worker/XlsxWorkerEditor';

export { XlsxSaveRefusedError } from './XlsxSaveRefusedError';
export type {
  XlsxEditorProps, XlsxDocumentChange, XlsxInThreadEditorProps, XlsxLegacyEditorProps, XlsxEditorApi,
  XlsxPointPosition, XlsxEditorCollaborationOptions, XlsxWorkerViewerApi, XlsxWorkerViewerProps,
  XlsxAsyncEditorProps, XlsxWorkerEditorProps,
} from './XlsxEditor.types';

function XlsxAsyncEditorSession(props: Omit<XlsxAsyncEditorProps, 'file'> & { file?: Uint8Array }) {
  const inThread = props.experimentalWorkerOpen === false;
  const identity = useMemo(() => ({}), [props.file, props.collaboration, inThread]);
  const start = () => ({ identity, reason: inThread ? 'deprecated' as const
    : typeof Worker !== 'function' ? 'no-worker' as const
    : typeof OffscreenCanvas !== 'function' ? 'no-offscreen-canvas' as const : undefined });
  const [choice, setChoice] = useState(start);
  const engine = choice.identity === identity ? choice : start();
  if (choice !== engine) setChoice(engine);
  useEffect(() => { if (engine.reason) warnXlsxEngine(engine.reason); }, [engine.reason]);
  if (!engine.reason) prefetchPeerlessXlsxEditor();
  return engine.reason ? <XlsxAsyncInThreadSessionLoader {...props} /> : <XlsxWorkerEditor {...props} backend="peerless" />;
}

/**
 * `<XlsxEditor />`: the editor shell. A dpr-aware canvas paints the grid; DOM
 * overlays (selection marquee, active-cell outline, in-cell editor) sit above it,
 * positioned from the same display-list geometry the painter uses. A top bar
 * holds the name box, formula bar, and save/undo/redo; an offscreen `role=grid`
 * mirror serves screen readers. All compute lives in `@betteroffice/xlsx`; this
 * layer is framework glue: keyboard/mouse wiring, focus flow, and DOM chrome.
 */
export function XlsxEditor(props: XlsxAsyncEditorProps): React.JSX.Element;
export function XlsxEditor(props: XlsxLegacyEditorProps): React.JSX.Element;
export function XlsxEditor(props: XlsxWorkerViewerProps): React.JSX.Element;
export function XlsxEditor(props: XlsxWorkerEditorProps): React.JSX.Element;
export function XlsxEditor(props: WorkerSessionProps): React.JSX.Element;
export function XlsxEditor(props: XlsxEditorProps): React.JSX.Element;
export function XlsxEditor(props: XlsxEditorProps | XlsxLegacyEditorProps | XlsxWorkerViewerProps | XlsxWorkerEditorProps | WorkerSessionProps): React.JSX.Element {
  const file = useMemo(() => props.file instanceof ArrayBuffer ? new Uint8Array(props.file) : props.file, [props.file]);
  const normalized = { ...props, file };
  return <LocaleProvider i18n={props.i18n}>
    {normalized[LEGACY_SYNC_API] === true ? <LegacyXlsxEditorLoader {...normalized} /> : <XlsxAsyncEditorSession {...normalized} />}
  </LocaleProvider>;
}
