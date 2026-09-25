import { createRoot } from 'react-dom/client';
import { useEffect, useRef, useState } from 'react';
import {
  DocxEditor,
  defineDocxPlugin,
  type DocxEditorRef,
  type DocxPluginContext,
} from '@betteroffice/docx-react';
import documentUrl from '../../../apps/demo/public/betteroffice-demo.docx?url';
import '../../../packages/docx-react/src/styles/editor.css';

interface ProbeState {
  version: string | null;
  paragraphs: { paraId: string; text: string }[];
}

type ProbeContext = DocxPluginContext<ProbeState>;

interface ProbeWindow {
  editor: DocxEditorRef | null;
  navigation: unknown;
  marks: number;
  reload(): Promise<void>;
}

const probe: ProbeWindow = {
  editor: null,
  navigation: null,
  marks: 0,
  async reload() {
    const buffer = await (await fetch(documentUrl)).arrayBuffer();
    await probe.editor!.loadDocumentBuffer(buffer);
  },
};
(window as unknown as { __probe: ProbeWindow }).__probe = probe;

async function refresh(context: ProbeContext) {
  const read = await context.read.readParagraphs({ view: 'accepted' });
  if (!read.ok) return;
  const paragraphs = read.paragraphs
    .filter((paragraph) => paragraph.text.trim().length > 0)
    .map(({ paraId, text }) => ({ paraId, text }));
  context.setState({ version: read.version, paragraphs }, read.version);
}

function ProbePanel({ context }: { context: ProbeContext }) {
  const last = context.state.paragraphs.at(-1);
  return (
    <div>
      <button
        type="button"
        data-testid="probe-navigate"
        disabled={!last}
        onClick={() =>
          void context.run(async (action) => {
            probe.navigation = await action.navigation.scrollToParagraph(
              { story: 'body', paraId: last!.paraId },
              { expectVersion: action.snapshot.version }
            );
          })
        }
      >
        Go to the last paragraph
      </button>
      <button
        type="button"
        data-testid="probe-append"
        onClick={() =>
          void context.run(async (action) => {
            const read = await action.read.readParagraphs({ view: 'accepted' });
            const first = read.ok ? read.paragraphs.find((p) => p.text.trim()) : undefined;
            if (!read.ok || !first || !action.edits) return;
            await action.edits.applyEdits({
              expectVersion: read.version,
              steps: [
                {
                  op: 'insertText',
                  target: { kind: 'paragraph', story: 'body', paraId: first.paraId },
                  at: 'end',
                  text: ' (checked)',
                },
              ],
            });
          })
        }
      >
        Append
      </button>
      <input data-testid="probe-input" aria-label="Probe note" />
    </div>
  );
}

const alignment = defineDocxPlugin<ProbeState>({
  id: 'probe.alignment',
  createState: () => ({ version: null, paragraphs: [] }),
  async onEvent(context, event) {
    if (event.type === 'load' || event.type === 'document-change') await refresh(context);
  },
  panel: { title: 'Probe', placement: 'right', render: ProbePanel },
  overlay: ({ geometry }) => (
    <>
      {Array.from({ length: Math.min(geometry.layout.pageCount, 2) }, (_, index) => {
        const bounds = geometry.dom.getPageBounds(index);
        if (!bounds) return null;
        const box = geometry.toOverlayRect(bounds);
        return (
          <div
            key={index}
            data-probe-page={index}
            style={{
              position: 'absolute',
              left: box.x,
              top: box.y,
              width: box.width,
              height: box.height,
              outline: '1px solid rgba(255, 0, 0, 0.4)',
            }}
          />
        );
      })}
    </>
  ),
  getSidebarItems(context) {
    const first = context.state.paragraphs[0];
    const version = context.state.version;
    return first && version
      ? [
          {
            id: 'first',
            anchor: { version, story: 'body', paraId: first.paraId },
            render: ({ isExpanded, onToggleExpand, measureRef }) => (
              <div ref={measureRef}>
                <button
                  type="button"
                  data-testid="probe-card"
                  data-version={version}
                  data-expanded={String(isExpanded)}
                  onClick={onToggleExpand}
                  style={{ display: 'block', width: '100%', height: 60 }}
                >
                  First paragraph
                </button>
              </div>
            ),
          },
        ]
      : [];
  },
  commands: [
    {
      id: 'mark',
      label: 'Probe mark',
      mutatesDocument: false,
      shortcuts: ['Mod+Alt+Shift+P'],
      execute() {
        probe.marks += 1;
        return { ok: true, status: 'executed' };
      },
    },
  ],
  toolbar: ['mark'],
});

const second = defineDocxPlugin({
  id: 'probe.second',
  createState: () => null,
  panel: {
    title: 'Second',
    placement: 'right',
    render: () => <p data-testid="second-panel">Second panel</p>,
  },
});

const PLUGINS = [alignment, second];
const GRANTS = { 'probe.alignment': { document: 'write', editBatches: true } } as const;

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  const editor = useRef<DocxEditorRef>(null);
  useEffect(() => {
    void fetch(documentUrl)
      .then((response) => response.arrayBuffer())
      .then(setBuffer);
  }, []);
  useEffect(() => {
    probe.editor = editor.current;
  });
  if (!buffer) return null;
  return (
    <div style={{ height: '100%' }}>
      <DocxEditor ref={editor} documentBuffer={buffer} plugins={PLUGINS} pluginGrants={GRANTS} />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
