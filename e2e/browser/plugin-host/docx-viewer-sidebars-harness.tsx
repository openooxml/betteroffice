import { createRoot } from 'react-dom/client';
import { useEffect, useRef, useState } from 'react';
import JSZip from 'jszip';
import { DocxEditor, defineDocxPlugin, type DocxEditorRef } from '@betteroffice/docx-react';
import { parseDocx, repackDocx } from '@betteroffice/docx/docx';
import type { Document } from '@betteroffice/docx/types/document';
import { ResidentEngineWorkerClient, type YrsSession } from '@betteroffice/docx/yrs';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import fixtureUrl from '../../../apps/demo/public/betteroffice-demo.docx?url';
import '../../../packages/docx-react/src/styles/editor.css';

const options = new URLSearchParams(window.location.search);
const noCopy = options.get('noCopy') === '1';
const viewing = options.get('kind') === 'viewing';
const publicFixture = options.get('fixture') === 'public';
const sourceKind = options.get('source');
const residentWorkers = new Set<Worker>();
const killedWorkers = new WeakSet<Worker>();
if (noCopy) {
  const postMessage = Worker.prototype.postMessage;
  Worker.prototype.postMessage = function (message: unknown, ...args: unknown[]) {
    if (message && typeof message === 'object' && 'type' in message && message.type === 'open') {
      residentWorkers.add(this);
    }
    return Reflect.apply(postMessage, this, [message, ...args]);
  };
}
setGoogleFontsEnabled(false);

async function viewerDocx(): Promise<ArrayBuffer> {
  const run = (text: string) => `<w:r><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr><w:t>${text}</w:t></w:r>`;
  const paragraph = (id: number, text: string, props = '') =>
    `<w:p w14:paraId="${id.toString(16).padStart(8, '0')}"><w:pPr>${props}</w:pPr>${run(text)}</w:p>`;
  const text = Array.from({ length: 25 }, (_, index) => `word${index}`).join(' ');
  const body =
    paragraph(1, 'First heading', '<w:pStyle w:val="Heading1"/>') +
    `<w:p w14:paraId="00000002"><w:commentRangeStart w:id="7"/>${run('Commented text on page one.')}<w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>` +
    `<w:p w14:paraId="00000003"><w:ins w:id="11" w:author="Document writer" w:date="2026-10-01T00:00:00Z">${run('Inserted text on page one.')}</w:ins></w:p>` +
    paragraph(4, text) + paragraph(5, text) +
    paragraph(6, 'Second heading', '<w:pStyle w:val="Heading2"/><w:pageBreakBefore/>') +
    paragraph(7, text) + paragraph(8, text) +
    paragraph(9, 'Page three heading', '<w:pStyle w:val="Heading1"/><w:pageBreakBefore/>') +
    paragraph(10, text);
  const zip = new JSZip();
  zip.file('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
  zip.file('_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/_rels/document.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('word/styles.xml',
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style></w:styles>');
  zip.file('word/comments.xml',
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="7" w:author="Document reader" w:date="2026-10-01T00:00:00Z"><w:p><w:r><w:t>Check this page one comment.</w:t></w:r></w:p></w:comment></w:comments>');
  zip.file('word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`);
  return zip.generateAsync({ type: 'arraybuffer' });
}

const methods = [
  'resolveComment', 'listRevisions', 'headings', 'paragraphs',
  'storySegments', 'locateParagraph', 'listComments', 'selection',
] as const;
const counts = Object.fromEntries(methods.map((method) => [method, 0])) as Record<typeof methods[number], number>;
const loadMethods = ['openDocx', 'openDocxPreview', 'loadState', 'applyUpdate'] as const;
type DocumentLoads = Record<typeof loadMethods[number], number>;
const documentLoads: DocumentLoads[] = [];
const loadEvents: { session: number; method: typeof loadMethods[number]; at: number }[] = [];
let documentWorker: ResidentEngineWorkerClient | null = null;
const documentRead = ResidentEngineWorkerClient.prototype.documentRead;
ResidentEngineWorkerClient.prototype.documentRead = function (this: ResidentEngineWorkerClient, ...args) {
  documentWorker = this;
  return documentRead.apply(this, args);
} as typeof documentRead;
const probe = {
  editor: null as DocxEditorRef | null,
  sessions: [] as YrsSession[],
  errors: [] as string[],
  reportedErrors: [] as Error[],
  copies: [] as string[],
  saveStarted: null as number | null,
  workersOpened() { return residentWorkers.size; },
  crashResidentWorker() {
    const worker = [...residentWorkers].reverse().find((entry) => !killedWorkers.has(entry));
    if (!worker) throw new Error('No resident worker is available to fail');
    killedWorkers.add(worker);
    worker.terminate();
    worker.dispatchEvent(new ErrorEvent('error', { message: 'Synthetic resident worker failure' }));
  },
  async saveForTest() {
    const before = this.mainDocumentLoads();
    this.saveStarted = performance.now();
    const bytes = await this.editor!.save();
    if (!bytes) throw new Error('Save returned no bytes');
    const zip = await JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml')?.async('string');
    return {
      before,
      saveStarted: this.saveStarted,
      byteLength: bytes.byteLength,
      signature: Array.from(new Uint8Array(bytes).slice(0, 4)),
      validDocument: !!xml?.includes('First heading'),
      after: this.mainDocumentLoads(),
    };
  },
  sidebarOpen: false,
  mainDocumentLoads() {
    return {
      sessionsCaptured: this.sessions.length,
      total: documentLoads.reduce((total, counts) =>
        total + loadMethods.reduce((sum, method) => sum + counts[method], 0), 0),
      sessions: documentLoads.map((counts) => ({ ...counts })),
      events: loadEvents.map((event) => ({ ...event })),
    };
  },
  replica() {
    const loaded = this.mainDocumentLoads().total > 0;
    return { started: loaded, loaded };
  },
  sessionReads() { return { ...counts }; },
  async commentAnchors(id: string) {
    if (!this.editor || !documentWorker) throw new Error('Viewer is not ready');
    const { version } = await this.editor.getProposals();
    const { value } = await documentWorker.documentRead({
      kind: 'sidebar', commentIds: [id], expectVersion: version,
    });
    return value?.comments.find((comment) => comment.id === id)?.anchors ?? null;
  },
};

export type ViewerSidebarsProbe = typeof probe;
(window as unknown as { __viewerSidebarsProbe: ViewerSidebarsProbe }).__viewerSidebarsProbe = probe;
(globalThis as unknown as {
  __workerProposalTest: { captureSession(session: YrsSession): void };
}).__workerProposalTest = {
  captureSession(session) {
    if (probe.sessions.includes(session)) return;
    const index = probe.sessions.length;
    probe.sessions.push(session);
    const loads: DocumentLoads = { openDocx: 0, openDocxPreview: 0, loadState: 0, applyUpdate: 0 };
    documentLoads.push(loads);
    const loadTarget = session as unknown as Record<string, (...args: unknown[]) => unknown>;
    for (const method of loadMethods) {
      const original = loadTarget[method]!;
      loadTarget[method] = (...args) => {
        loads[method] += 1;
        loadEvents.push({ session: index, method, at: performance.now() });
        return original.apply(session, args);
      };
    }
    const target = session as unknown as Record<string, (...args: unknown[]) => unknown>;
    for (const method of methods) {
      const original = target[method]!;
      target[method] = (...args) => {
        counts[method] += 1;
        return original.apply(session, args);
      };
    }
  },
};

window.addEventListener('copy', (event) => {
  probe.copies.push(event.clipboardData?.getData('text/plain') ?? '');
});

const viewerCard = defineDocxPlugin<{ anchor: { version: string; story: string; paraId: string } | null }>({
  id: 'probe.viewer-no-copy',
  createState: () => ({ anchor: null }),
  async onEvent(context, event) {
    if (event.type !== 'load' && event.type !== 'layout-change') return;
    const result = await context.read.findText({
      text: 'Commented text', within: { kind: 'story', story: 'body' }, view: 'accepted',
    });
    if (!result.ok) throw new Error(result.failure.message);
    const match = result.matches[0];
    if (match) context.setState({
      anchor: { version: result.version, story: match.range.story, paraId: match.range.start.paraId },
    }, result.version);
  },
  getSidebarItems: (context) => context.state.anchor ? [{
    id: 'viewer-card',
    anchor: context.state.anchor,
    render: ({ measureRef }) => <div ref={measureRef} data-testid="viewer-plugin-card">Viewer plugin card</div>,
  }] : [],
});
const viewerPlugins = [viewerCard];

const fonts = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };
const faces = [{ family: 'Liberation Sans', src: fontUrl }];

function Harness() {
  const [source, setSource] = useState<{ buffer?: ArrayBuffer; document?: Document } | null>(null);
  const pendingDocument = useRef<Document | null>(null);
  const editor = useRef<DocxEditorRef>(null);
  useEffect(() => {
    void (async () => {
      if (!publicFixture) {
        setSource({ buffer: await viewerDocx() });
        return;
      }
      const document = await parseDocx(await (await fetch(fixtureUrl)).arrayBuffer(), { preloadFonts: false });
      const paragraph = document.package.document.content.find((block) => block.type === 'paragraph');
      if (!paragraph) throw new Error('The public fixture has no body paragraph');
      paragraph.content.push({ type: 'run', content: [{ type: 'text', text: ' Parsed document host edit.' }] });
      if (sourceKind === 'prop') setSource({ document });
      else if (sourceKind === 'loadDocument') {
        pendingDocument.current = document;
        setSource({});
      } else setSource({ buffer: await repackDocx(document) });
    })().catch((error: Error) => {
      probe.errors.push(error.message);
      probe.reportedErrors.push(error);
    });
  }, []);
  useEffect(() => {
    probe.editor = editor.current;
    if (editor.current && pendingDocument.current) {
      const document = pendingDocument.current;
      pendingDocument.current = null;
      editor.current.loadDocument(document);
    }
  });
  if (!source) return null;
  return (
    <div style={{ height: '100%' }}>
      <DocxEditor
        ref={editor}
        documentBuffer={source.buffer}
        document={source.document}
        experimentalWorkerOpen
        readOnly={!viewing}
        mode={viewing ? 'viewing' : undefined}
        plugins={noCopy && !publicFixture ? viewerPlugins : undefined}
        allowHostProposals
        onFirstPagePainted={() => {
          if (source.document) setSource({});
        }}
        onCommentsSidebarOpenChange={(open) => { probe.sidebarOpen = open; }}
        fonts={faces}
        measurementFontProvider={fonts}
        onError={(error) => {
          probe.errors.push(error.message);
          probe.reportedErrors.push(error);
        }}
        onPluginError={noCopy ? (error) => probe.errors.push(String(error.error)) : undefined}
      />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
