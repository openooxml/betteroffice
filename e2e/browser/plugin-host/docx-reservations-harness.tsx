import { createRoot } from 'react-dom/client';
import { createRef, useEffect, useState } from 'react';
import JSZip from 'jszip';
import {
  DocxEditor,
  defineDocxPlugin,
  type DocxEditorRef,
  type DocxPluginContext,
  type DocxPluginGeometry,
  type DocxReservation,
} from '@betteroffice/docx-react';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

setGoogleFontsEnabled(false);
const fonts = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };
const faces = [{ family: 'Liberation Sans', src: fontUrl }];

async function reservationDocument(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  const body = ['First paragraph text', 'Second paragraph text'].map((text, index) =>
    `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}"><w:r>` +
    '<w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr>' +
    `<w:t>${text}</w:t></w:r></w:p>`).join('');
  zip.file('word/document.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    `xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}` +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'arraybuffer' });
}

interface Target {
  version: string;
  paraId: string;
  text: string;
}

type Context = DocxPluginContext<Target | null>;

interface ReservationProbe {
  editor: DocxEditorRef | null;
  geometry: DocxPluginGeometry | null;
  target: Target | null;
  selectionChanges: number;
  setPlugins(enabled: boolean): void;
  setReadOnly(readOnly: boolean): void;
}

const editorRef = createRef<DocxEditorRef>();
const probe: ReservationProbe = {
  get editor() {
    return editorRef.current;
  },
  geometry: null,
  target: null,
  selectionChanges: 0,
  setPlugins() {},
  setReadOnly() {},
};
(window as unknown as { __reservations: ReservationProbe }).__reservations = probe;

async function refresh(context: Context) {
  const read = await context.read.readParagraphs({ view: 'accepted' });
  if (!read.ok) return;
  const first = read.paragraphs.find((paragraph) => /^\S+ \S+/.test(paragraph.text));
  if (first) {
    context.setState({ version: read.version, paraId: first.paraId, text: first.text }, read.version);
  }
}

function items(target: Target): DocxReservation[] {
  const at = (offset: number) => ({ paraId: target.paraId, offset });
  const range = (from: number, to: number) => ({
    kind: 'range' as const,
    version: target.version,
    range: { story: 'body', start: at(from), end: at(to), view: 'accepted' as const },
  });
  const word = target.text.indexOf(' ') + 1;
  return [
    { id: 'chip', kind: 'inline', width: 48, height: 24, anchor: range(word, word) },
    { id: 'gap', kind: 'block', side: 'after', height: 80, anchor: range(0, 0) },
  ];
}

function Overlay({ context, geometry }: { context: Context; geometry: DocxPluginGeometry }) {
  useEffect(() => {
    probe.geometry = geometry;
    probe.target = context.state;
  }, [context, geometry]);
  const target = context.state;
  if (!target || geometry.layout.reservationsVersion !== target.version) return null;
  return (
    <>
      {['chip', 'gap'].map((id) => {
        const result = geometry.getAnchorGeometry({ kind: 'reservation', id });
        if (!result.ok) return null;
        const { x, y, width, height } = result.anchor;
        return (
          <div
            key={id}
            data-reserved={id}
            style={{
              position: 'absolute',
              left: x,
              top: y,
              width,
              height,
              background: 'rgba(0, 120, 255, 0.25)',
            }}
          />
        );
      })}
    </>
  );
}

const reserve = defineDocxPlugin<Target | null>({
  id: 'probe.reserve',
  createState: () => null,
  async onEvent(context, event) {
    if (event.type === 'selection-change') probe.selectionChanges++;
    if (event.type === 'load' || event.type === 'document-change') await refresh(context);
  },
  reservations: (context) =>
    context.state && context.state.version === context.snapshot.version
      ? { version: context.snapshot.version, items: items(context.state) }
      : null,
  overlay: Overlay,
});

const PLUGINS = [reserve];
const NONE: typeof PLUGINS = [];

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  const [enabled, setEnabled] = useState(true);
  const [readOnly, setReadOnly] = useState(new URLSearchParams(location.search).has('readOnly'));
  useEffect(() => {
    void reservationDocument().then(setBuffer);
  }, []);
  useEffect(() => {
    probe.setPlugins = setEnabled;
    probe.setReadOnly = setReadOnly;
  });
  if (!buffer) return null;
  return (
    <div style={{ height: '100%' }}>
      <DocxEditor
        ref={editorRef}
        documentBuffer={buffer}
        experimentalPeerless
        plugins={enabled ? PLUGINS : NONE}
        readOnly={readOnly}
        fonts={faces}
        measurementFontProvider={fonts}
      />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
