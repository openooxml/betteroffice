import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import {
  XlsxEditor,
  defineXlsxPlugin,
  type XlsxCommandStore,
  type XlsxEditorProps,
  type XlsxShortcuts,
  type ShortcutIssue,
} from '@betteroffice/xlsx-react';
import workbookUrl from '../../../packages/xlsx/test-fixtures/sample.xlsx?url';

const mode = new URLSearchParams(location.search).get('mode') ?? 'async';
const route = (
  mode === 'legacy' ? { legacySyncApi: true, experimentalWorkerOpen: false }
  : mode === 'legacy-worker' ? { legacySyncApi: true, experimentalWorkerOpen: true }
  : {}
) as Partial<XlsxEditorProps>;
let open: (state: { file: Uint8Array; shortcuts: XlsxShortcuts }) => void = () => {};
const probe = {
  commands: null as XlsxCommandStore | null,
  issues: [] as ShortcutIssue[],
  hostKeys: [] as { key: string; prevented: boolean }[],
  pluginRuns: 0,
  async load(shortcuts: XlsxShortcuts) {
    open({ file: new Uint8Array(await (await fetch(workbookUrl)).arrayBuffer()), shortcuts });
  },
  bold() { return probe.commands?.getState('bold').active ?? null; },
  italic() { return probe.commands?.getState('italic').active ?? null; },
};
(window as unknown as { __shortcuts: typeof probe }).__shortcuts = probe;
window.addEventListener('keydown', (event) => {
  if (!event.ctrlKey && !event.metaKey) return;
  probe.hostKeys.push({ key: event.key.toLowerCase(), prevented: event.defaultPrevented });
});
const plugin = defineXlsxPlugin({
  id: 'acme.keys',
  createState: () => null,
  commands: [{
    id: 'run', label: 'Run', mutatesDocument: false, shortcuts: ['Mod+Shift+K'],
    execute: () => { probe.pluginRuns += 1; return { ok: true, status: 'noop' }; },
  }],
} as never);
const plugins = [plugin] as never;

function Harness() {
  const [state, setState] = useState<{ file: Uint8Array; shortcuts: XlsxShortcuts } | null>(null);
  open = setState;
  if (!state) return null;
  return <XlsxEditor {...(route as object)} file={state.file} plugins={plugins} shortcuts={state.shortcuts}
    onShortcutIssue={(issue) => probe.issues.push(issue)}
    onReady={(api) => { probe.commands = api.commands; }} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
