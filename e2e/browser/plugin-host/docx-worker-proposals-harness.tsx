import { createRoot } from 'react-dom/client';
import { useEffect, useRef, useState } from 'react';
import {
  DocxEditor,
  defineDocxPlugin,
  type DocxEditorRef,
  type DocxPluginContext,
  type DocxPluginGeometry,
} from '@betteroffice/docx-react';
import type { DocxProposalSnapshot, YrsSession } from '@betteroffice/docx/yrs';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import { pagedDocx } from '../../../packages/docx-react/src/components/DocxEditor/__fixtures__/pagedDocx';
import { workerOpenReplicaPending } from '../../../packages/docx-react/src/components/DocxEditor/internals/workerOpenReplica';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

interface OverlayState {
  snapshot: DocxProposalSnapshot | null;
  eventSerial: number;
}

type OverlayContext = DocxPluginContext<OverlayState>;
type NavigationResult = Awaited<ReturnType<OverlayContext['navigation']['scrollToParagraph']>>;
interface ViewState {
  scrollTop: number;
  selection: ReturnType<YrsSession['selection']>;
}
const options = new URLSearchParams(window.location.search);
const readOnly = options.get('readOnly') !== 'false';
const revisions = options.get('revisions') === '1';
const keepSidebarClosed = options.get('sidebar') === 'closed';
const reportChanges = options.get('onChange') === '1';
setGoogleFontsEnabled(false);

const loadMethods = ['openDocx', 'openDocxPreview', 'loadState', 'applyUpdate'] as const;
type DocumentLoads = Record<typeof loadMethods[number], number>;
const sessions: YrsSession[] = [];
const documentLoads: DocumentLoads[] = [];
const probe = {
  editor: null as DocxEditorRef | null,
  session: null as YrsSession | null,
  context: null as OverlayContext | null,
  captures: 0,
  hydratedBeforeSidebar: false,
  sidebarOpened: false,
  sidebarOpen: false,
  sidebarOpenChanges: [] as boolean[],
  beforeSidebarOpen: null as ViewState | null,
  layoutComplete: null as number | null,
  renderedDomContextCalls: 0,
  contentChanges: [] as { bodyContainsProposedText: boolean }[],
  documentChanges: [] as string[],
  load: null as { version: string; sessionVersion: string; snapshotVersion: string } | null,
  events: { load: 0, 'proposal-change': 0, 'layout-change': 0 },
  eventSerial: 0,
  errors: [] as string[],
  mainDocumentLoads() {
    return {
      sessionsCaptured: sessions.length,
      total: documentLoads.reduce((total, counts) =>
        total + loadMethods.reduce((sum, method) => sum + counts[method], 0), 0),
      sessions: documentLoads.map((counts) => ({ ...counts })),
    };
  },
  status() {
    return {
      pending: this.session ? workerOpenReplicaPending(this.session) : null,
      captures: this.captures,
      mainDocumentLoads: this.mainDocumentLoads(),
      hydratedBeforeSidebar: this.hydratedBeforeSidebar,
      sidebarOpen: this.sidebarOpen,
      sidebarOpenChanges: [...this.sidebarOpenChanges],
      beforeSidebarOpen: this.beforeSidebarOpen,
      layoutComplete: this.layoutComplete,
      renderedDomContextCalls: this.renderedDomContextCalls,
      contentChanges: [...this.contentChanges],
      documentChanges: [...this.documentChanges],
      load: this.load,
      events: { ...this.events },
      errors: [...this.errors],
    };
  },
  view(): ViewState {
    const scroller = document.querySelector<HTMLElement>('.docx-editor__scroll-container')!;
    return {
      scrollTop: scroller.scrollTop,
      selection: this.mainDocumentLoads().total === 0 ? null : this.session!.selection(),
    };
  },
  async navigate(
    target: { story: string; paraId: string },
    expectVersion: string
  ): Promise<NavigationResult | null> {
    let result: NavigationResult | null = null;
    await this.context!.run(async (context) => {
      result = await context.navigation.scrollToParagraph(target, { expectVersion });
    });
    return result;
  },
  async toggleSidebar() {
    this.sidebarOpened = true;
    return this.editor!.commands.execute('commentsSidebar', null);
  },
};

export type WorkerProposalProbe = typeof probe;

(window as unknown as { __workerProposalProbe: WorkerProposalProbe }).__workerProposalProbe = probe;
(globalThis as unknown as {
  __workerProposalTest: { captureSession(session: YrsSession): void };
}).__workerProposalTest = {
  captureSession(session) {
    if (sessions.includes(session)) return;
    sessions.push(session);
    probe.session = session;
    probe.captures += 1;
    const loads: DocumentLoads = { openDocx: 0, openDocxPreview: 0, loadState: 0, applyUpdate: 0 };
    documentLoads.push(loads);
    const target = session as unknown as Record<string, (...args: unknown[]) => unknown>;
    for (const method of loadMethods) {
      const original = target[method]!;
      target[method] = (...args) => {
        loads[method] += 1;
        if (!probe.sidebarOpened) probe.hydratedBeforeSidebar = true;
        return original.apply(session, args);
      };
    }
  },
};

function ProposalOverlay({
  context,
  geometry,
}: {
  context: OverlayContext;
  geometry: DocxPluginGeometry;
}) {
  useEffect(() => {
    probe.context = context;
  }, [context]);
  const snapshot = context.state.snapshot;
  return (
    <div
      data-testid="proposal-overlay"
      data-version={snapshot?.version}
      data-preview-version={snapshot?.previewVersion}
      data-event-serial={context.state.eventSerial}
      data-layout-version={geometry.layout.version}
      data-layout-preview-version={geometry.layout.previewVersion}
    >
      {snapshot?.proposals.map((proposal) => {
        const result = geometry.getAnchorGeometry({ kind: 'proposal', id: proposal.id });
        return (
          <div
            key={proposal.id}
            data-proposal-id={proposal.id}
            data-state={proposal.state}
            data-ok={String(result.ok)}
            data-failure={result.ok ? '' : result.failure.code}
            style={
              result.ok
                ? {
                    position: 'absolute',
                    left: result.anchor.x,
                    top: result.anchor.y,
                    width: 2,
                    height: Math.max(2, result.anchor.height),
                    background: 'purple',
                  }
                : undefined
            }
          />
        );
      })}
    </div>
  );
}

const overlay = defineDocxPlugin<OverlayState>({
  id: 'probe.worker-proposals',
  createState: () => ({ snapshot: null, eventSerial: 0 }),
  async onEvent(context, event) {
    if (event.type !== 'load' && event.type !== 'proposal-change' && event.type !== 'layout-change') {
      return;
    }
    probe.events[event.type] += 1;
    probe.eventSerial += 1;
    if (event.type === 'load') {
      if (probe.events.load === 1) {
        void probe.editor!.whenLayoutComplete().then(
          () => { probe.layoutComplete = performance.now(); },
          (error: Error) => probe.errors.push(error.message)
        );
      }
      const version = await context.read.version();
      if (!version.ok) throw new Error(version.failure.message);
      probe.load = {
        version: version.version,
        sessionVersion: probe.session!.version(),
        snapshotVersion: context.snapshot.version,
      };
    }
    const snapshot = probe.session!.getProposals();
    context.setState({ snapshot, eventSerial: probe.eventSerial }, snapshot.version);
  },
  overlay: ProposalOverlay,
});

const PLUGINS = [overlay];
const fonts = {
  resolve: () => async () => (await fetch(fontUrl)).arrayBuffer(),
};
const faces = [{ family: 'Liberation Sans', src: fontUrl }];

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const editor = useRef<DocxEditorRef>(null);
  useEffect(() => {
    void pagedDocx(6, 12, { trackedInsertion: revisions }).then(setBuffer);
  }, []);
  useEffect(() => {
    probe.editor = editor.current;
    probe.sidebarOpen = sidebarOpen;
  });
  if (!buffer) return null;
  return (
    <div style={{ height: '100%' }}>
      <DocxEditor
        ref={editor}
        documentBuffer={buffer}
        experimentalWorkerOpen
        readOnly={readOnly}
        allowHostProposals
        commentsSidebarOpen={keepSidebarClosed ? false : sidebarOpen}
        onCommentsSidebarOpenChange={(open) => {
          probe.sidebarOpenChanges.push(open);
          if (!keepSidebarClosed) {
            if (open) {
              probe.beforeSidebarOpen ??= probe.view();
              probe.sidebarOpened = true;
            }
            setSidebarOpen(open);
          }
        }}
        showHostProposalsInSidebar={sidebarOpen}
        plugins={PLUGINS}
        fonts={faces}
        measurementFontProvider={fonts}
        onRenderedDomContextReady={() => { probe.renderedDomContextCalls += 1; }}
        onChange={
          reportChanges ? (document) => {
            probe.contentChanges.push({
              bodyContainsProposedText:
                JSON.stringify(document.package.document).includes('Reviewed 1'),
            });
          } : undefined
        }
        onDocumentChange={reportChanges ? ({ version }) => { probe.documentChanges.push(version); } : undefined}
        onError={(error) => probe.errors.push(error.message)}
        onPluginError={(error) => probe.errors.push(String(error.error))}
      />
    </div>
  );
}

createRoot(document.getElementById('root')!).render(<Harness />);
