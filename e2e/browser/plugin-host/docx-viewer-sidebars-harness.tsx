import { createRoot } from 'react-dom/client';
import { useEffect, useRef, useState } from 'react';
import JSZip from 'jszip';
import { DocxEditor, type DocxEditorRef } from '@betteroffice/docx-react';
import { ResidentEngineWorkerClient, type YrsSession } from '@betteroffice/docx/yrs';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import {
  workerOpenReplicaPending,
  workerOpenReplicaStarted,
} from '../../../packages/docx-react/src/components/DocxEditor/internals/workerOpenReplica';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

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
  sidebarOpen: false,
  replica() {
    return {
      started: this.sessions.some((session) => workerOpenReplicaStarted(session)),
      loaded: this.sessions.some((session) => !workerOpenReplicaPending(session)),
    };
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
    probe.sessions.push(session);
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

const fonts = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };
const faces = [{ family: 'Liberation Sans', src: fontUrl }];

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  const editor = useRef<DocxEditorRef>(null);
  useEffect(() => { void viewerDocx().then(setBuffer); }, []);
  useEffect(() => { probe.editor = editor.current; });
  if (!buffer) return null;
  return (
    <div style={{ height: '100%' }}>
      <DocxEditor
        ref={editor}
        documentBuffer={buffer}
        experimentalWorkerOpen
        readOnly
        allowHostProposals
        onCommentsSidebarOpenChange={(open) => { probe.sidebarOpen = open; }}
        fonts={faces}
        measurementFontProvider={fonts}
        onError={(error) => probe.errors.push(error.message)}
      />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
