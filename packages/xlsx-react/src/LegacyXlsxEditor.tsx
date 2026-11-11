import { useEffect, useMemo, useRef, useState } from 'react';
import type {
  GeneralEditorSessionProps, WorkerSessionProps, XlsxInThreadEditorProps, XlsxLegacyEditorProps,
  XlsxWorkerEditorProps, XlsxWorkerViewerProps,
} from './XlsxEditor.types';
import { XlsxEditorContent } from './XlsxEditorContent';
import { warnXlsxEngine, type XlsxEngineWarning } from './engineWarnings';
import type { XlsxPluginError, XlsxPluginFailureCode } from './plugins/types';
import { XlsxSessionViewer } from './viewer/XlsxSessionViewer';
import { LegacyXlsxWorkerEditor } from './worker/LegacyXlsxWorkerEditor';
import type { EditableSessionWorkbookProps } from './worker/useEditableSessionWorkbook';

export function LegacyXlsxEditor(props: GeneralEditorSessionProps | XlsxLegacyEditorProps | XlsxWorkerViewerProps | XlsxWorkerEditorProps | WorkerSessionProps): React.JSX.Element {
  useEffect(() => { warnXlsxEngine('legacy-sync-api'); }, []);
  const workerOpen = props.experimentalWorkerOpen;
  const editable = !props.readOnly || ('allowHostProposals' in props && props.allowHostProposals === true);
  const collaboration = props.collaboration;
  const identity = useMemo(() => ({}), [props.file, collaboration, workerOpen]);
  const start = () => ({ identity, reason: workerOpen === false ? 'deprecated' as XlsxEngineWarning
    : workerOpen !== true ? 'in-thread' as const
    : typeof Worker !== 'function' ? 'no-worker' as XlsxEngineWarning : undefined, editor: editable });
  const [choice, setChoice] = useState(start);
  const session = choice.identity === identity ? choice : start();
  const editor = session.editor || editable;
  if (choice !== session || choice.editor !== editor) setChoice({ ...session, editor });
  useEffect(() => { if (session.reason && session.reason !== 'in-thread') warnXlsxEngine(session.reason); }, [session.reason]);
  const viewer = !session.reason && !editor;
  const reportedPlugins = useRef({ identity, keys: new Set<string>() });
  useEffect(() => {
    if (!viewer || reportedPlugins.current.identity !== identity) reportedPlugins.current = { identity, keys: new Set() };
    if (!viewer) return;
    const plugins = (props.plugins ?? []).map((plugin) => ({ plugin, key: `${plugin.id}\u0000${plugin.revision === undefined ? '' : `${typeof plugin.revision}:${plugin.revision}`}` }));
    const reported = reportedPlugins.current.keys;
    const present = new Set(plugins.map(({ key }) => key));
    for (const key of reported) if (!present.has(key)) reported.delete(key);
    for (const { plugin, key } of plugins) {
      if (reported.has(key)) continue;
      reported.add(key);
      const error: XlsxPluginError = {
        pluginId: plugin.id,
        generation: null,
        phase: 'definition',
        error: Object.assign(new Error('Plugins are not supported in the deprecated read-only worker viewer; omit the legacy sync API to use the async editor.'), {
          code: 'unsupported-policy' satisfies XlsxPluginFailureCode,
        }),
      };
      if (!props.onPluginError) {
        console.error(`[XlsxEditor] plugin "${error.pluginId}" failed (${error.phase})`, error.error);
        continue;
      }
      try {
        props.onPluginError(error);
      } catch (reporterError) {
        console.error('[XlsxEditor] onPluginError threw', reporterError);
      }
    }
  }, [viewer, identity, props.plugins, props.onPluginError]);
  return session.reason ? <XlsxEditorContent {...props as Omit<XlsxInThreadEditorProps, 'file'> & { file?: Uint8Array }} />
    : editor ? <LegacyXlsxWorkerEditor {...props as EditableSessionWorkbookProps} />
    : <XlsxSessionViewer {...props as XlsxWorkerViewerProps & { file?: Uint8Array }} />;
}
