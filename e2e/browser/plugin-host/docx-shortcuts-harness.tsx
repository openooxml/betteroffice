import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import { DocxEditor, type DocxEditorProps, type DocxEditorRef, type DocxShortcuts, type ShortcutIssue } from '@betteroffice/docx-react';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

setGoogleFontsEnabled(false);
const worker = new URLSearchParams(location.search).get('mode') !== 'legacy';
const route: Partial<DocxEditorProps> = worker ? {} : { legacySyncApi: true, experimentalWorkerOpen: false } as never;
let open: (buffer: ArrayBuffer, shortcuts: DocxShortcuts) => void = () => {};
const probe = {
  editor: null as DocxEditorRef | null,
  issues: [] as ShortcutIssue[],
  hostKeys: [] as { key: string; prevented: boolean }[],
  swallow: null as string | null,
  load(bytes: number[], shortcuts: DocxShortcuts) { open(Uint8Array.from(bytes).buffer, shortcuts); },
  async saved() { return [...new Uint8Array((await this.editor!.save())!)]; },
};
(window as unknown as { __shortcuts: typeof probe }).__shortcuts = probe;
window.addEventListener('keydown', (event) => {
  if (!event.ctrlKey && !event.metaKey) return;
  if (event.key.toLowerCase() === probe.swallow) event.preventDefault();
  probe.hostKeys.push({ key: event.key.toLowerCase(), prevented: event.defaultPrevented });
});
const fonts = [{ family: 'Liberation Sans', src: fontUrl }];
const measurementFontProvider = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };

function Harness() {
  const [state, setState] = useState<{ buffer: ArrayBuffer; shortcuts: DocxShortcuts } | null>(null);
  open = (buffer, shortcuts) => setState({ buffer, shortcuts });
  return state && <DocxEditor {...route} ref={(editor) => { probe.editor = editor; }} documentBuffer={state.buffer}
    fonts={fonts} measurementFontProvider={measurementFontProvider} shortcuts={state.shortcuts}
    onShortcutIssue={(issue) => probe.issues.push(issue)} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
