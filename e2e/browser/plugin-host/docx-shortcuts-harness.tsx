import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import { DocxEditor, defineDocxPlugin, type DocxEditorRef, type DocxShortcuts, type ShortcutIssue } from '@betteroffice/docx-react';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

setGoogleFontsEnabled(false);
const worker = new URLSearchParams(location.search).get('mode') !== 'legacy';
const second = new URLSearchParams(location.search).has('two');
const route = (worker ? {} : { legacySyncApi: true, experimentalWorkerOpen: false }) as Record<never, never>;
let open: (buffer: ArrayBuffer, shortcuts: DocxShortcuts) => void = () => {};
const probe = {
  editor: null as DocxEditorRef | null,
  issues: [] as ShortcutIssue[],
  hostKeys: [] as { key: string; prevented: boolean }[],
  swallow: null as string | null,
  pluginRuns: 0,
  errors: [] as string[],
  load(bytes: number[], shortcuts: DocxShortcuts) { open(Uint8Array.from(bytes).buffer, shortcuts); },
  async saved() { return [...new Uint8Array((await this.editor!.save())!)]; },
};
(window as unknown as { __shortcuts: typeof probe }).__shortcuts = probe;
window.addEventListener('keydown', (event) => {
  if (!event.ctrlKey && !event.metaKey) return;
  if (event.key.toLowerCase() === probe.swallow) event.preventDefault();
  probe.hostKeys.push({ key: event.key.toLowerCase(), prevented: event.defaultPrevented });
});
const plugin = defineDocxPlugin({
  id: 'acme.keys',
  createState: () => null,
  commands: [{
    id: 'run', label: 'Run', mutatesDocument: false, shortcuts: ['Mod+K'],
    execute: () => { probe.pluginRuns += 1; return { ok: true, status: 'noop' }; },
  }],
} as never);
const plugins = [plugin] as never;
const fonts = [{ family: 'Liberation Sans', src: fontUrl }];
const measurementFontProvider = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };

function Harness() {
  const [state, setState] = useState<{ buffer: ArrayBuffer; shortcuts: DocxShortcuts } | null>(null);
  open = (buffer, shortcuts) => setState({ buffer, shortcuts });
  return state && <>
    <div style={{ height: second ? '50%' : '100%' }}><DocxEditor {...route} ref={(editor) => { probe.editor = editor; }} documentBuffer={state.buffer}
      fonts={fonts} measurementFontProvider={measurementFontProvider} shortcuts={state.shortcuts} plugins={plugins}
      onShortcutIssue={(issue) => probe.issues.push(issue)} onError={(error) => probe.errors.push(error.message)} /></div>
    {second && <div style={{ height: '50%' }}><DocxEditor {...route} documentBuffer={state.buffer.slice(0)} fonts={fonts} measurementFontProvider={measurementFontProvider} /></div>}
  </>;
}
createRoot(document.getElementById('root')!).render(<Harness />);
