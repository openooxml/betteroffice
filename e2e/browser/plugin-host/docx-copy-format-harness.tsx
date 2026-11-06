import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import { DocxEditor, type DocxEditorProps, type DocxEditorRef } from '@betteroffice/docx-react';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

setGoogleFontsEnabled(false);
const params = new URLSearchParams(location.search);
const route = (params.get('mode') === 'legacy' ? { legacySyncApi: true, experimentalWorkerOpen: false } : {}) as Partial<DocxEditorProps>;
const editorMode = params.get('suggesting') ? { mode: 'suggesting' as const } : {};
let open: (buffer: ArrayBuffer) => void = () => {};
const probe = {
  editor: null as DocxEditorRef | null,
  load(bytes: number[]) { open(Uint8Array.from(bytes).buffer); },
  async saved() { return [...new Uint8Array((await this.editor!.save())!)]; },
};
(window as unknown as { __copyFormat: typeof probe }).__copyFormat = probe;
const fonts = [{ family: 'Liberation Sans', src: fontUrl }];
const measurementFontProvider = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };

function Harness() {
  const [state, setState] = useState<{ buffer: ArrayBuffer; key: number } | null>(null);
  open = (buffer) => setState((previous) => ({ buffer, key: (previous?.key ?? 0) + 1 }));
  return state && <DocxEditor key={state.key} {...route} {...editorMode} ref={(editor) => { probe.editor = editor; }}
    documentBuffer={state.buffer} fonts={fonts} measurementFontProvider={measurementFontProvider} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
