import { createRoot } from 'react-dom/client';
import { useEffect, useState } from 'react';
import { DocxEditor, defineDocxPlugin, type DocxEditorRef, type DocxPluginGeometry } from '@betteroffice/docx-react';
import { DOCX_MINI_TOOLBAR_ID, docxMiniToolbar } from '@betteroffice/docx-react/mini-toolbar';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

setGoogleFontsEnabled(false);
const params = new URLSearchParams(location.search);
let open: (buffer: ArrayBuffer) => void = () => {};
let setReadOnly: (value: boolean) => void = () => {};
const probe = {
  editor: null as DocxEditorRef | null,
  geometry: null as DocxPluginGeometry | null,
  load(bytes: number[]) { open(Uint8Array.from(bytes).buffer); },
  readOnly(value: boolean) { setReadOnly(value); },
  async saved() { return [...new Uint8Array((await this.editor!.save())!)]; },
};
(window as unknown as { __miniToolbar: typeof probe }).__miniToolbar = probe;
const fonts = [{ family: 'Liberation Sans', src: fontUrl }];
const measurementFontProvider = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };

function GeometryProbe({ geometry }: { geometry: DocxPluginGeometry }) {
  useEffect(() => { probe.geometry = geometry; }, [geometry]);
  return null;
}

const geometryProbe = defineDocxPlugin({ id: 'probe.geometry', createState: () => null, overlay: GeometryProbe });
const PLUGINS = [geometryProbe, docxMiniToolbar()];
const GRANTS = {
  [DOCX_MINI_TOOLBAR_ID]: {
    document: 'write',
    commands: ['bold', 'italic', 'underline', 'strikethrough', 'clearFormatting', 'pasteAsPlainText'],
  },
} as const;
const editorMode = params.get('mode') === 'viewing' ? { mode: 'viewing' as const } : {};

function Harness() {
  const [state, setState] = useState<{ buffer: ArrayBuffer; key: number } | null>(null);
  const [readOnly, changeReadOnly] = useState(params.get('readonly') === '1');
  open = (buffer) => setState((previous) => ({ buffer, key: (previous?.key ?? 0) + 1 }));
  setReadOnly = changeReadOnly;
  return state && <><DocxEditor key={state.key} ref={(editor) => { probe.editor = editor; }} documentBuffer={state.buffer}
    fonts={fonts} measurementFontProvider={measurementFontProvider} plugins={PLUGINS} pluginGrants={GRANTS}
    readOnly={readOnly} {...editorMode} />
    <button type="button" style={{ position: 'fixed', right: 0, bottom: 0 }}>After the editor</button></>;
}
createRoot(document.getElementById('root')!).render(<Harness />);
