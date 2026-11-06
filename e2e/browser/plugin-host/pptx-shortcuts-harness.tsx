import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import {
  PptxEditor,
  definePptxPlugin,
  type PptxAsyncEditorApi,
  type PptxEditorApi,
  type PptxShortcuts,
  type ShortcutIssue,
} from '@betteroffice/pptx-react';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';

const legacy = new URLSearchParams(location.search).get('mode') === 'legacy';
const route = legacy ? { legacySyncApi: true, experimentalWorkerOpen: false } : {};

type Fonts = { family: string; bytes: Uint8Array }[];
let open: (state: { file: Uint8Array; fonts: Fonts; shortcuts: PptxShortcuts }) => void = () => {};
const probe = {
  editor: null as PptxEditorApi | PptxAsyncEditorApi | null,
  issues: [] as ShortcutIssue[],
  hostKeys: [] as { key: string; prevented: boolean }[],
  saves: 0,
  pluginRuns: 0,
  async load(bytes: number[], shortcuts: PptxShortcuts) {
    const fonts = [{ family: 'Liberation Sans', bytes: new Uint8Array(await (await fetch(fontUrl)).arrayBuffer()) }];
    open({ file: Uint8Array.from(bytes), fonts, shortcuts });
  },
  state(id: 'bold' | 'italic') { return probe.editor?.commands.getState(id).active ?? null; },
  async texts() {
    const read = await probe.editor!.readContent();
    return (read as unknown as { stories: { text: string }[] }).stories.map((story) => story.text);
  },
};
(window as unknown as { __shortcuts: typeof probe }).__shortcuts = probe;
window.addEventListener('keydown', (event) => {
  if (!event.ctrlKey && !event.metaKey) return;
  probe.hostKeys.push({ key: event.key.toLowerCase(), prevented: event.defaultPrevented });
});
const plugin = definePptxPlugin({
  id: 'acme.keys',
  createState: () => null,
  commands: [{
    id: 'run', label: 'Run', mutatesDocument: false, shortcuts: ['Mod+Shift+K'],
    execute: () => { probe.pluginRuns += 1; return { ok: true, status: 'noop' }; },
  }],
} as never);
const plugins = [plugin] as never;

function Harness() {
  const [state, setState] = useState<{ file: Uint8Array; fonts: Fonts; shortcuts: PptxShortcuts } | null>(null);
  open = setState;
  if (!state) return null;
  return <PptxEditor {...(route as object)} file={state.file} fonts={state.fonts} plugins={plugins}
    shortcuts={state.shortcuts} onShortcutIssue={(issue) => probe.issues.push(issue)}
    onSave={() => { probe.saves += 1; }} onReady={(api: PptxEditorApi | PptxAsyncEditorApi) => { probe.editor = api; }} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
