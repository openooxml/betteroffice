import { createRoot } from 'react-dom/client';
import { useEffect, useState } from 'react';
import { DocxEditor, defineDocxPlugin, type DocxEditorRef, type DocxPluginContext, type DocxPluginGeometry } from '@betteroffice/docx-react';
import type { DocxProposalInput, DocxProposalResult, DocxProposalSnapshot, DocxProposalState } from '@betteroffice/docx/yrs';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import { pagedDocx } from '../../../packages/docx-react/src/components/DocxEditor/__fixtures__/pagedDocx';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

type Paragraph = Awaited<ReturnType<DocxEditorRef['getParagraphIdentities']>>['paragraphs'][number];
interface State { snapshot: DocxProposalSnapshot; anchors: Record<'body' | 'header' | 'footer', Paragraph | undefined> | null }
type Context = DocxPluginContext<State | null>;
const params = new URLSearchParams(location.search);
const pages = Number(params.get('pages') ?? 120);
const viewing = params.get('viewing') !== '0';
const proposeAtOpen = params.get('propose') !== '0';
const pageTotal = params.get('pageTotal') === '1';
const headerFooter = params.get('headerFooter') === '1';
const firstProposalPage = Number(params.get('at') ?? 0);
const withFonts = params.get('fonts') !== '0';
const PER_PAGE = 12;
const STRIDE = 3;
const started = performance.now();
setGoogleFontsEnabled(false);
const proposals: DocxProposalInput[] = Array.from({ length: 10 }, (_, index) => ({
  id: `host-${index + 1}`, op: 'replaceText', search: 'Original', replaceWith: `Reviewed ${index + 1}`,
  paragraph: { kind: 'persisted', story: { partUri: '/word/document.xml', kind: 'body' },
    paraId: (firstProposalPage * PER_PAGE + index * STRIDE + 1).toString(16).padStart(8, '0') },
  suggest: { author: 'Host reviewer', date: '2026-10-01T00:00:00Z' },
}));
const elapsed = () => performance.now() - started;
const probe = {
  editor: null as DocxEditorRef | null,
  context: null as Context | null,
  geometry: null as DocxPluginGeometry | null,
  firstPagePaintedCalls: 0,
  started,
  proposed: null as DocxProposalResult | null,
  timings: { firstPagePainted: null as number | null, firstInk: null as number | null, preview: null as number | null,
    load: null as number | null, layoutComplete: null as number | null },
  errors: [] as string[],
  history: [] as { id: string; state: DocxProposalState }[][],
  async decide(ids: string[], state: DocxProposalState) {
    const current = await this.editor!.getProposals();
    const result = await this.editor!.setProposalStates({ expectVersion: current.version, expectPreviewVersion: current.previewVersion,
      changes: ids.map(id => ({ id, state })) });
    if (result.ok) this.history.push(current.proposals.filter(proposal => ids.includes(proposal.id)).map(({ id, state }) => ({ id, state })));
    return result;
  },
  async undoDecision() {
    const current = await this.editor!.getProposals();
    return this.editor!.setProposalStates({ expectVersion: current.version, expectPreviewVersion: current.previewVersion,
      changes: this.history.pop()! });
  },
  async next(index: number) {
    const snapshot = await this.editor!.getProposals();
    let result: Awaited<ReturnType<Context['navigation']['scrollToParagraph']>> | null = null;
    await this.context!.run(async context => {
      result = await context.navigation.scrollToParagraph(snapshot.proposals[index]!.paragraph, { expectVersion: snapshot.version });
    });
    return result;
  },
  async jumpFirst() {
    const snapshot = this.context!.state!.snapshot;
    const placed = snapshot.proposals.flatMap(proposal => {
      const anchor = this.geometry?.getAnchorGeometry({ kind: 'proposal', id: proposal.id });
      return anchor?.ok ? [{ proposal, anchor: anchor.anchor }] : [];
    }).sort((a, b) => a.anchor.pageIndex - b.anchor.pageIndex || a.anchor.y - b.anchor.y);
    const first = placed[0];
    if (!first) return null;
    let result: Awaited<ReturnType<Context['navigation']['scrollToParagraph']>> | null = null;
    await this.context!.run(async context => {
      result = await context.navigation.scrollToParagraph(first.proposal.paragraph, { expectVersion: snapshot.version });
    });
    return { id: first.proposal.id, pageIndex: first.anchor.pageIndex, result };
  },
  pageInked(pageIndex: number) {
    const canvas = document.querySelector<HTMLCanvasElement>(`.canvas-page[data-page-index="${pageIndex}"] canvas`);
    return !!canvas && canvas.width > 0 && inked(canvas);
  },
  firstPage() {
    const canvas = document.querySelector<HTMLCanvasElement>('[data-peerless-docx] [data-page-index="0"] canvas');
    return canvas && canvas.width > 0 ? canvas.toDataURL() : null;
  },
};
export type ViewerHostProposalProbe = typeof probe;
(window as unknown as { __viewerHostProposalProbe: ViewerHostProposalProbe }).__viewerHostProposalProbe = probe;

const scratch = document.createElement('canvas');
const inked = (canvas: HTMLCanvasElement) => {
  scratch.width = canvas.width;
  scratch.height = Math.min(canvas.height, 400);
  const context = scratch.getContext('2d', { willReadFrequently: true })!;
  context.fillStyle = 'white';
  context.fillRect(0, 0, scratch.width, scratch.height);
  context.drawImage(canvas, 0, 0);
  const pixels = context.getImageData(0, 0, scratch.width, scratch.height).data;
  for (let index = 0; index < pixels.length; index += 4) if (pixels[index]! < 128) return true;
  return false;
};
const sample = () => {
  if (document.querySelector('[aria-label="Document preview"] canvas')) probe.timings.preview ??= elapsed();
  const canvas = document.querySelector<HTMLCanvasElement>('[data-peerless-docx] [data-page-index="0"] canvas');
  if (canvas && canvas.width > 0 && inked(canvas)) probe.timings.firstInk = elapsed();
  else requestAnimationFrame(sample);
};
requestAnimationFrame(sample);

function Anchors({ anchors, geometry }: { anchors: NonNullable<State['anchors']>; geometry: DocxPluginGeometry }) {
  const { body, header, footer } = anchors;
  const targets: [string, Parameters<DocxPluginGeometry['getAnchorGeometry']>[0] | null][] = [
    ['paragraph', body?.persisted ? { kind: 'paragraph', paragraph: body.persisted } : null],
    ['search', body?.persisted ? { kind: 'search', paragraph: body.persisted, text: 'paragraph 1' } : null],
    ['range', body?.session ? { kind: 'range', version: geometry.layout.version, range: { story: body.session.story,
      start: { paraId: body.session.paraId, offset: 0 }, end: { paraId: body.session.paraId, offset: 4 }, view: 'accepted' } } : null],
    ['header', header?.persisted ? { kind: 'paragraph', paragraph: header.persisted } : null],
    ['footer', footer?.persisted ? { kind: 'paragraph', paragraph: footer.persisted } : null],
  ];
  return <>{targets.map(([name, target]) => {
    const placed = target ? geometry.getAnchorGeometry(target) : null;
    return <div key={name} data-anchor-target={name} data-ok={String(placed?.ok === true)}
      data-failure={!placed ? 'no-anchor' : placed.ok ? '' : placed.failure.code} data-rects={placed?.ok ? placed.rects.length : 0} />;
  })}</>;
}

function Overlay({ context, geometry }: { context: Context; geometry: DocxPluginGeometry }) {
  useEffect(() => { probe.context = context; }, [context]);
  useEffect(() => { probe.geometry = geometry; }, [geometry]);
  return <div data-testid="host-proposals" data-preview={context.state?.snapshot.previewVersion} data-layout-preview={geometry.layout.previewVersion}>
    {context.state?.anchors && <Anchors anchors={context.state.anchors} geometry={geometry} />}
    {context.state?.snapshot.proposals.map((proposal, index) => {
      const placed = geometry.getAnchorGeometry({ kind: 'proposal', id: proposal.id });
      const text = geometry.getAnchorGeometry({ kind: 'search', paragraph: proposal.paragraph, text: `Reviewed ${index + 1}` });
      const aligned = placed.ok && text.ok && placed.anchor.pageIndex === text.anchor.pageIndex &&
        Math.abs(placed.anchor.x - text.anchor.x) < 1 && Math.abs(placed.anchor.y - text.anchor.y) < 1;
      return <div key={proposal.id} data-proposal-id={proposal.id} data-state={proposal.state} data-ok={String(placed.ok)}
        data-text-aligned={String(aligned)} data-page={placed.ok ? placed.anchor.pageIndex : ''}
        data-in-page={String(placed.ok && placed.anchor.y >= placed.pageRect.y && placed.anchor.y <= placed.pageRect.y + placed.pageRect.height)}
        data-failure={placed.ok ? '' : placed.failure.code}
        style={placed.ok ? { position: 'absolute', left: placed.anchor.x, top: placed.anchor.y,
          width: 4, height: Math.max(4, placed.anchor.height), background: 'purple' } : undefined} />;
    })}
  </div>;
}
let anchors: State['anchors'] = null;
const plugins = [defineDocxPlugin<State | null>({ id: 'host-proposals', createState: () => null,
  async onEvent(context, event) {
    if (event.type === 'load') probe.timings.load ??= elapsed();
    if (event.type !== 'load' && event.type !== 'proposal-change' && event.type !== 'layout-change') return;
    if (headerFooter && !anchors) {
      const { paragraphs } = await probe.editor!.getParagraphIdentities();
      const find = (kind: 'body' | 'header' | 'footer') => paragraphs.find(paragraph => paragraph.persisted?.story.kind === kind);
      anchors = { body: find('body'), header: find('header'), footer: find('footer') };
    }
    const snapshot = await probe.editor!.getProposals();
    context.setState({ snapshot, anchors }, snapshot.version);
  }, overlay: Overlay,
})];
const fonts = [{ family: 'Liberation Sans', src: fontUrl }];
const measurementFontProvider = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };
let opened = false;
const attach = (editor: DocxEditorRef | null) => {
  probe.editor = editor;
  if (!editor || opened) return;
  opened = true;
  void editor.whenLayoutComplete({ timeoutMs: 170_000 }).then(() => { probe.timings.layoutComplete = elapsed(); },
    (error: Error) => probe.errors.push(error.message));
  if (proposeAtOpen) void (async () => {
    const current = await editor.getProposals();
    probe.proposed = await editor.proposeChanges({ expectVersion: current.version, proposals });
  })().catch((error: unknown) => probe.errors.push(error instanceof Error ? error.message : String(error)));
};

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  useEffect(() => { void pagedDocx(pages, PER_PAGE, { pageTotal, headerFooter }).then(setBuffer); }, []);
  return buffer && <div style={{ height: '100%' }}>
    <DocxEditor ref={attach} documentBuffer={buffer} mode={viewing ? 'viewing' : 'editing'} readOnly={viewing} allowHostProposals
      previewFirstPage showToolbar={false} plugins={plugins} fonts={withFonts ? fonts : undefined} measurementFontProvider={measurementFontProvider}
      onFirstPagePainted={() => { probe.firstPagePaintedCalls += 1; probe.timings.firstPagePainted ??= elapsed(); }}
      onError={error => probe.errors.push(error.message)} onPluginError={error => probe.errors.push(String(error.error))} />
  </div>;
}
createRoot(document.getElementById('root')!).render(<Harness />);
