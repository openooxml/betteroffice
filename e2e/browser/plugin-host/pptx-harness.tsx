import { createRoot } from 'react-dom/client';
import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import {
  PptxEditor,
  definePptxPlugin,
  type PptxEditorApi,
  type PptxAsyncEditorApi,
  type PptxPluginContext,
  type PptxPluginEvent,
  type PptxPluginGeometry,
  type PptxPluginNavigationResult,
  type PptxPluginSelection,
} from '@betteroffice/pptx-react';
import deckUrl from '../../../apps/demo/public/betteroffice-demo.pptx?url';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';

interface ProbeState {
  slides: { slideId: string; shapeId: string | null }[];
}

type ProbeContext = PptxPluginContext<ProbeState>;

interface ProbeWindow {
  mutation: unknown;
  runMutation(): Promise<unknown>;
  events: PptxPluginEvent[];
  editor: PptxEditorApi | PptxAsyncEditorApi | null;
  navigation: PptxPluginNavigationResult | null;
  selection: PptxPluginSelection;
  geometry: PptxPluginGeometry | null;
  marks: number;
}

const workerOnly = new URLSearchParams(location.search).get('engine') === 'worker';
const probe: ProbeWindow = { events: [], editor: null, navigation: null, selection: null, geometry: null, marks: 0, mutation: null, runMutation: async () => null };
(window as unknown as { __probe: ProbeWindow }).__probe = probe;

async function refresh(context: ProbeContext) {
  const read = await context.read.readContent();
  if (!read.ok) return;
  const slides = read.slides.map((slide) => ({
    slideId: slide.id,
    shapeId: read.stories.find((story) => story.slideId === slide.id)?.shapeId ?? null,
  }));
  context.setState({ slides }, read.version);
}

function ProbePanel({ context }: { context: ProbeContext }) {
  probe.runMutation = () => context.commands.execute('bold', null);
  const last = context.state.slides.at(-1);
  return (
    <div>
      <button type="button" data-testid="probe-mutation" data-version={context.snapshot.version} onMouseDown={(event) => event.preventDefault()}
        onClick={() => { void probe.runMutation().then((result) => { probe.mutation = result; }); }}>
        Apply formatting
      </button>
      <button type="button" data-testid="probe-timer-mutation" onMouseDown={(event) => event.preventDefault()}
        onClick={async () => {
          probe.mutation = null;
          await new Promise((resolve) => setTimeout(resolve, 0));
          probe.mutation = await probe.runMutation();
        }}>
        Apply formatting after a timer
      </button>
      <button
        type="button"
        data-testid="probe-navigate"
        disabled={!last}
        onClick={() =>
          void context.run(async (action) => {
            probe.navigation = await action.navigation.goToSlide(
              { slideId: last!.slideId },
              { expectVersion: action.snapshot.version }
            );
          })
        }
      >
        Go to the last slide
      </button>
      <input data-testid="probe-input" aria-label="Probe note" />
      {createPortal(
        <input
          data-testid="portal-input"
          aria-label="Portal note"
          style={{ position: 'fixed', right: 8, bottom: 8 }}
        />,
        document.body
      )}
    </div>
  );
}

const alignment = definePptxPlugin<ProbeState>({
  id: 'probe.alignment',
  createState: () => ({ slides: [] }),
  async onEvent(context, event) {
    probe.events.push(event);
    if (event.type === 'selection-change') probe.selection = event.selection;
    if (event.type === 'load' || event.type === 'document-change') await refresh(context);
  },
  panel: { title: 'Probe', placement: 'right', render: ProbePanel },
  overlay: ({ context, geometry }) => {
    probe.geometry = geometry;
    const slide = geometry.toOverlayRect({
      space: 'slide-px',
      rect: { x: 0, y: 0, width: geometry.layout.width, height: geometry.layout.height },
    });
    const shapeId = context.state.slides.find(
      (entry) => entry.slideId === geometry.layout.slideId
    )?.shapeId;
    const shape = shapeId ? geometry.getShapeRect(shapeId) : null;
    const selected = context.snapshot.selection;
    const target = selected?.target;
    const anchors = selected && target?.kind === 'text' ? [
      { name: 'selection', answer: geometry.getAnchorGeometry({ kind: 'selection' }) },
      { name: 'range', answer: geometry.getAnchorGeometry({ kind: 'range', version: geometry.layout.version,
        slideId: selected.slideId, shapeId: target.shapeId, storyId: target.storyId, anchor: target.anchor, focus: target.focus }) },
    ] : [];

    return (
      <>
        {anchors.map(({ name, answer }) => answer.ok && <div key={name} data-probe-anchor={name}
          style={{ position: 'absolute', left: answer.anchor.x, top: answer.anchor.y,
            width: Math.max(1, answer.anchor.width), height: answer.anchor.height }} />)}
        {slide && (
          <input
            data-testid="overlay-input"
            aria-label="Overlay note"
            style={{
              position: 'absolute',
              left: slide.x + 8,
              top: slide.y + 8,
              width: 120,
              pointerEvents: 'auto',
            }}
          />
        )}
        {slide && (
          <div
            data-probe-slide={geometry.layout.slide}
            data-zoom={geometry.layout.zoom}
            style={{
              position: 'absolute',
              left: slide.x,
              top: slide.y,
              width: slide.width,
              height: slide.height,
              outline: '1px solid rgba(255, 0, 0, 0.4)',
            }}
          />
        )}
        {shape && (
          <div
            data-probe-shape={shapeId}
            style={{
              position: 'absolute',
              left: shape.x,
              top: shape.y,
              width: shape.width,
              height: shape.height,
              background: 'rgba(0, 128, 255, 0.08)',
            }}
          />
        )}
      </>
    );
  },
  commands: [
    {
      id: 'mark',
      label: 'Probe mark',
      mutatesDocument: false,
      shortcuts: ['Mod+Shift+K'],
      execute() {
        probe.marks += 1;
        return { ok: true, status: 'executed' };
      },
    },
  ],
  toolbar: ['mark'],
});

const second = definePptxPlugin({
  id: 'probe.second',
  createState: () => null,
  panel: {
    title: 'Second',
    placement: 'right',
    render: () => <p data-testid="second-panel">Second panel</p>,
  },
});

const PLUGINS = [alignment, second];

function Harness() {
  const [file, setFile] = useState<Uint8Array | null>(null);
  const [fonts, setFonts] = useState<{ family: string; bytes: Uint8Array }[] | null>(null);
  useEffect(() => {
    void Promise.all([fetch(deckUrl), fetch(fontUrl)])
      .then((responses) => Promise.all(responses.map((response) => response.arrayBuffer())))
      .then(([deck, font]) => {
        setFonts([{ family: 'Liberation Sans', bytes: new Uint8Array(font) }]);
        setFile(new Uint8Array(deck));
      });
  }, []);
  if (!file || !fonts) return null;
  return (
    <div style={{ height: '100%' }}>
      <PptxEditor {...(workerOnly ? {} : { legacySyncApi: true, experimentalWorkerOpen: false }) as Record<never, never>}
        file={file}
        fonts={fonts}
        plugins={PLUGINS}
        pluginGrants={{ 'probe.alignment': { document: 'write', commands: ['bold'] } }}
        onReady={(api) => {
          probe.editor = api;
        }}
      />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
