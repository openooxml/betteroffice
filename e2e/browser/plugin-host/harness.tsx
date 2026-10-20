import { createFontProvider } from '@betteroffice/fonts';
import { createRoot } from 'react-dom/client';
import { useCallback, useEffect, useRef, useState } from 'react';
import {
  DocxEditor,
  defineDocxPlugin,
  type DocxEditorRef,
  type DocxPluginContext,
  type DocxPluginGeometry,
  type DocxPluginSidebarItem,
} from '@betteroffice/docx-react';
import documentUrl from '../../../apps/demo/public/betteroffice-demo.docx?url';
import { repeatedHeaderDocx } from '../../../packages/docx-react/src/plugins/__fixtures__/storyAnchorsDocx';
import '../../../packages/docx-react/src/styles/editor.css';

interface ProbeState {
  version: string | null;
  paragraphs: { paraId: string; text: string }[];
}

type ProbeContext = DocxPluginContext<ProbeState>;

interface ProbeWindow {
  editor: DocxEditorRef | null;
  geometry: DocxPluginGeometry | null;
  navigation: unknown;
  marks: number;
  reload(): Promise<void>;
}

const fonts = createFontProvider();
const repeatedHeader = new URLSearchParams(location.search).get('doc') === 'repeated-header';

const probe: ProbeWindow = {
  editor: null,
  geometry: null,
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
        data-testid="probe-navigate-focus"
        disabled={!last}
        onClick={() =>
          void context.run(async (action) => {
            probe.navigation = await action.navigation.scrollToParagraph(
              { story: 'body', paraId: last!.paraId },
              { expectVersion: action.snapshot.version, focus: true }
            );
          })
        }
      >
        Edit the last paragraph
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

/** Rects of `{{header}}` in the repeated header, one element per rect. */
function HeaderProbe({ geometry }: { geometry: DocxPluginGeometry }) {
  const result = geometry.getAnchorGeometry({
    kind: 'range',
    version: geometry.layout.version,
    range: {
      story: 'hf:rIdH',
      start: { paraId: '10000001', offset: 9 },
      end: { paraId: '10000001', offset: 19 },
      view: 'accepted',
    },
  });
  if (!result.ok) return <div data-probe-header-failure={result.failure.code} hidden />;
  return (
    <>
      {result.rects.map((rect, index) => (
        <div
          key={index}
          data-probe-header-rect={rect.pageIndex}
          style={{ position: 'absolute', left: rect.x, top: rect.y, width: rect.width, height: rect.height, outline: '1px solid blue' }}
        />
      ))}
    </>
  );
}

function GeometryProbe({ geometry }: { geometry: DocxPluginGeometry }) {
  useEffect(() => {
    probe.geometry = geometry;
  }, [geometry]);
  return null;
}

function ProbeCard({
  item,
  isExpanded,
  onToggleExpand,
  measureRef,
}: {
  item: DocxPluginSidebarItem<ProbeState>;
  isExpanded: boolean;
  onToggleExpand(): void;
  measureRef(element: HTMLDivElement | null): void;
}) {
  const [clicks, setClicks] = useState(0);
  return (
    <div ref={measureRef}>
      <button
        type="button"
        data-testid="probe-card"
        data-version={item.anchor.version}
        data-expanded={String(isExpanded)}
        data-clicks={clicks}
        onClick={() => {
          setClicks((count) => count + 1);
          onToggleExpand();
        }}
        style={{ display: 'block', width: '100%', height: 60 }}
      >
        First paragraph
      </button>
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
      <GeometryProbe geometry={geometry} />
      {repeatedHeader && <HeaderProbe geometry={geometry} />}
      {Array.from({ length: geometry.layout.pageCount }, (_, index) => {
        const bounds = geometry.dom.getPageBounds(index);
        const box = bounds && geometry.toOverlayRect(bounds);
        if (!box) return null;
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
            render: ProbeCard,
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
  const setEditor = useCallback((value: DocxEditorRef | null) => { editor.current = value; probe.editor = value; }, []);
  useEffect(() => {
    void (repeatedHeader
      ? repeatedHeaderDocx().then((bytes) => new Uint8Array(bytes).buffer)
      : fetch(documentUrl).then((response) => response.arrayBuffer())
    ).then(setBuffer);
  }, []);
  if (!buffer) return null;
  return (
    <div style={{ height: '100%' }}>
      <DocxEditor ref={setEditor} documentBuffer={buffer} plugins={PLUGINS} pluginGrants={GRANTS} measurementFontProvider={new URLSearchParams(location.search).get('peerless') === '1' ? fonts : undefined} {...(new URLSearchParams(location.search).get('peerless') === '1' ? {} : { legacySyncApi: true as const })} />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
