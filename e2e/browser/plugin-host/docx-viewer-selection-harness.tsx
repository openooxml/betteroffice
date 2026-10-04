import { createRoot } from 'react-dom/client';
import { useEffect, useRef, useState } from 'react';
import JSZip from 'jszip';
import { DocxEditor, type DocxEditorRef } from '@betteroffice/docx-react';
import type { YrsSession } from '@betteroffice/docx/yrs';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const options = new URLSearchParams(window.location.search);
const preview = options.get('preview') === '1';
setGoogleFontsEnabled(false);

/** Paragraph `index` (from 1): long enough to wrap over several lines. */
export function paragraphText(index: number): string {
  return `P${index} ${Array.from({ length: 90 }, (_, word) => `w${index}x${word}`).join(' ')}.`;
}

const PARAGRAPHS = 12;

async function viewerDocx(): Promise<ArrayBuffer> {
  const run = '<w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr>';
  const body = Array.from({ length: PARAGRAPHS }, (_, index) => {
    const id = (index + 1).toString(16).padStart(8, '0');
    return `<w:p w14:paraId="${id}"><w:r>${run}<w:t>${paragraphText(index + 1)}</w:t></w:r></w:p>`;
  }).join('');
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>'
  );
  zip.file(
    'word/document.xml',
    `<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${body}</w:body></w:document>`
  );
  return zip.generateAsync({ type: 'arraybuffer' });
}

const loadMethods = ['openDocx', 'openDocxPreview', 'loadState', 'applyUpdate'] as const;
type DocumentLoads = Record<typeof loadMethods[number], number>;
const documentLoads: DocumentLoads[] = [];
const loadEvents: { session: number; method: typeof loadMethods[number]; at: number }[] = [];
const probe = {
  editor: null as DocxEditorRef | null,
  sessions: [] as YrsSession[],
  copies: [] as string[],
  errors: [] as string[],
  paragraphText,
  paragraphs: PARAGRAPHS,
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
};

export type ViewerSelectionProbe = typeof probe;
(window as unknown as { __viewerSelectionProbe: ViewerSelectionProbe }).__viewerSelectionProbe = probe;
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
  },
};
window.addEventListener('copy', (event) => {
  probe.copies.push(event.clipboardData?.getData('text/plain') ?? '');
});

const fonts = {
  resolve: () => async () => (await fetch(fontUrl)).arrayBuffer(),
};
const faces = [{ family: 'Liberation Sans', src: fontUrl }];

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  const editor = useRef<DocxEditorRef>(null);
  useEffect(() => {
    void viewerDocx().then(setBuffer);
  }, []);
  useEffect(() => {
    probe.editor = editor.current;
  });
  if (!buffer) return null;
  return (
    <div style={{ height: '100%' }}>
      <DocxEditor
        ref={editor}
        documentBuffer={buffer}
        experimentalWorkerOpen
        previewFirstPage={preview}
        readOnly
        allowHostProposals
        fonts={faces}
        measurementFontProvider={fonts}
        onError={(error) => probe.errors.push(error.message)}
      />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
