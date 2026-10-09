import { createRoot } from 'react-dom/client';
import { useEffect, useState } from 'react';
import JSZip from 'jszip';
import {
  DocxEditor,
  defineDocxPlugin,
  type DocxEditorRef,
  type DocxGeometryTarget,
  type DocxPluginContext,
} from '@betteroffice/docx-react';
import type { DocxSourceParagraphAnchor, YrsSession } from '@betteroffice/docx/yrs';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import { workerOpenReplicaPending } from '../../../packages/docx-react/src/components/DocxEditor/internals/workerOpenReplica';
import documentUrl from '../../../apps/demo/public/betteroffice-demo.docx?url';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

type GeometryContext = DocxPluginContext<{ eventSerial: number }>;
const worker = new URLSearchParams(window.location.search).get('worker') === '1';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
setGoogleFontsEnabled(false);

const sessions = new Set<YrsSession>();
const probe = {
  editor: null as DocxEditorRef | null,
  session: null as YrsSession | null,
  context: null as GeometryContext | null,
  layoutComplete: null as number | null,
  events: { load: 0, 'proposal-change': 0, 'layout-change': 0 },
  errors: [] as string[],
  mainDocumentLoads: 0,
  sourceParagraphs: [] as { text: string; anchor: DocxSourceParagraphAnchor }[],
  pending() {
    return this.session ? workerOpenReplicaPending(this.session) : null;
  },
  anchorGeometry(targets: DocxGeometryTarget[]) {
    return targets.map((target) => this.context?.geometry?.getAnchorGeometry(target) ?? null);
  },
  async readParagraphs() {
    let result: Awaited<ReturnType<GeometryContext['read']['readParagraphs']>> | null = null;
    await this.context!.run(async (context) => {
      result = await context.read.readParagraphs({ story: 'body', view: 'accepted' });
    });
    const read = result as Awaited<ReturnType<GeometryContext['read']['readParagraphs']>> | null;
    if (!read?.ok) throw new Error(read?.failure.message ?? 'Paragraph read did not complete');
    this.session ??= this.editor!.getEditorRef()?.getYrsSession() ?? null;
    return {
      ...read,
      paragraphs: read.paragraphs.map((paragraph) => {
        const matches = this.sourceParagraphs.filter(({ text }) => text === paragraph.text);
        return { ...paragraph, anchor: matches.length === 1 ? matches[0].anchor : null };
      }),
    };
  },
};

(window as unknown as { __workerGeometryProbe: typeof probe }).__workerGeometryProbe = probe;
(globalThis as unknown as {
  __workerProposalTest: { captureSession(session: YrsSession): void };
}).__workerProposalTest = {
  captureSession(session) {
    if (sessions.has(session)) return;
    sessions.add(session);
    probe.session = session;
    const target = session as unknown as Record<string, (...args: unknown[]) => unknown>;
    for (const method of ['openDocx', 'openDocxPreview', 'loadState', 'applyUpdate']) {
      const original = target[method]!;
      target[method] = (...args) => {
        probe.mainDocumentLoads += 1;
        return original.apply(session, args);
      };
    }
  },
};

function GeometryOverlay({ context }: { context: GeometryContext }) {
  useEffect(() => {
    probe.context = context;
  }, [context]);
  return <div data-testid="geometry-overlay" data-event-serial={context.state.eventSerial} />;
}

const plugin = defineDocxPlugin<{ eventSerial: number }>({
  id: 'probe.worker-geometry',
  createState: () => ({ eventSerial: 0 }),
  onEvent(context, event) {
    if (event.type !== 'load' && event.type !== 'proposal-change' && event.type !== 'layout-change') {
      return;
    }
    probe.events[event.type] += 1;
    if (event.type === 'load' && probe.events.load === 1) {
      void probe.editor!.whenLayoutComplete().then(
        (pages) => { probe.layoutComplete = pages; },
        (error: Error) => probe.errors.push(error.message)
      );
    }
    context.setState({ eventSerial: context.state.eventSerial + 1 });
  },
  overlay: GeometryOverlay,
});

const PLUGINS = [plugin];
const fonts = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };
const faces = [{ family: 'Liberation Sans', src: fontUrl }];

async function loadDocument() {
  const response = await fetch(documentUrl);
  if (!response.ok) throw new Error(`Document fetch failed: ${response.status}`);
  const buffer = await response.arrayBuffer();
  const packageSha256 = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', buffer)))
    .map((byte) => byte.toString(16).padStart(2, '0')).join('');
  const zip = await JSZip.loadAsync(buffer);
  const xml = new DOMParser().parseFromString(await zip.file('word/document.xml')!.async('text'), 'text/xml');
  probe.sourceParagraphs = Array.from(xml.getElementsByTagNameNS(W, 'p')).flatMap((paragraph, paragraphOrdinal) =>
    paragraph.parentElement?.localName === 'body'
      ? [{
          text: Array.from(paragraph.getElementsByTagNameNS(W, 't')).map((text) => text.textContent ?? '').join(''),
          anchor: { kind: 'source' as const, packageSha256, partUri: '/word/document.xml', paragraphOrdinal },
        }]
      : []
  );
  return buffer;
}

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  useEffect(() => {
    void loadDocument().then(setBuffer, (error: Error) => probe.errors.push(error.message));
  }, []);
  if (!buffer) return null;
  return (
    <div style={{ height: '100%' }}>
      <DocxEditor
        ref={(editor) => { probe.editor = editor; }}
        documentBuffer={buffer}
        experimentalWorkerOpen={worker}
        readOnly
        allowHostProposals
        commentsSidebarOpen={false}
        showHostProposalsInSidebar={false}
        plugins={PLUGINS}
        fonts={faces}
        measurementFontProvider={fonts}
        onError={(error) => probe.errors.push(error.message)}
        onPluginError={(error) => probe.errors.push(String(error.error))}
      />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
