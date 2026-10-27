import { createRoot } from 'react-dom/client';
import { useEffect, useState } from 'react';
import { DocxEditor, defineDocxPlugin, type DocxEditorRef, type DocxPluginContext, type DocxPluginGeometry } from '@betteroffice/docx-react';
import { exportDocxMarkdown } from '@betteroffice/docx';
import type { DocxProposalSnapshot, DocxProposalState } from '@betteroffice/docx/yrs';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import { pagedDocx } from '../../../packages/docx-react/src/components/DocxEditor/__fixtures__/pagedDocx';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

type Context = DocxPluginContext<DocxProposalSnapshot | null>;
const arm = new URLSearchParams(location.search).get('arm') ?? 'worker';
setGoogleFontsEnabled(false);
const probe = {
  editor: null as DocxEditorRef | null,
  context: null as Context | null,
  events: [] as { generation: string; version: string; previewVersion: number }[],
  errors: [] as string[],
  history: [] as { id: string; state: DocxProposalState }[][],
  async propose() {
    const editor = this.editor!;
    const identities = await editor.getParagraphIdentities();
    const body = identities.paragraphs.filter(paragraph => paragraph.persisted?.story.kind === 'body');
    const anchors = [body[1]!, body[3]!, body[13]!, body[15]!].map(paragraph => paragraph.persisted!);
    const resolved = await editor.resolveParagraphAnchors(anchors);
    return editor.proposeChanges({ expectVersion: resolved.version, proposals: anchors.map((paragraph, index) => ({
      id: `change-${index}`, paragraph, op: 'replaceText', search: 'Original', replaceWith: `Reviewed ${index + 1}`,
      suggest: { author: 'Reviewer', date: '2026-10-01T00:00:00Z' },
    })) });
  },
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
  async next(index: number, focus = false) {
    const snapshot = await this.editor!.getProposals();
    const proposal = snapshot.proposals[index]!;
    let result: Awaited<ReturnType<Context['navigation']['scrollToParagraph']>> | null = null;
    await this.context!.run(async context => {
      result = await context.navigation.scrollToParagraph(proposal.paragraph, { expectVersion: snapshot.version, focus });
    });
    return result;
  },
  async savedText() {
    const bytes = await this.editor!.save();
    if (!bytes) throw new Error('No saved document');
    return (await exportDocxMarkdown(new Uint8Array(bytes), { revisionView: 'accepted', stories: ['body'] })).markdown;
  },
};
export type PeerlessProposalProbe = typeof probe;
(window as unknown as { __peerlessProposalProbe: PeerlessProposalProbe }).__peerlessProposalProbe = probe;

function Overlay({ context, geometry }: { context: Context; geometry: DocxPluginGeometry }) {
  useEffect(() => { probe.context = context; }, [context]);
  return <div data-testid="host-proposals" data-preview={context.state?.previewVersion} data-layout-preview={geometry.layout.previewVersion}>
    {context.state?.proposals.map((proposal, index) => {
      const placed = geometry.getAnchorGeometry({ kind: 'proposal', id: proposal.id });
      const text = geometry.getAnchorGeometry({ kind: 'search', paragraph: proposal.paragraph, text: `Reviewed ${index + 1}` });
      const aligned = placed.ok && text.ok && placed.anchor.pageIndex === text.anchor.pageIndex &&
        Math.abs(placed.anchor.x - text.anchor.x) < 1 && Math.abs(placed.anchor.y - text.anchor.y) < 1;
      return <div key={proposal.id} data-proposal-id={proposal.id} data-state={proposal.state} data-ok={String(placed.ok)}
        data-text-aligned={String(aligned)}
        data-page={placed.ok ? placed.anchor.pageIndex : ''} data-layout-version={geometry.layout.version}
        style={placed.ok ? { position: 'absolute', left: placed.anchor.x, top: placed.anchor.y,
          width: 4, height: Math.max(4, placed.anchor.height), background: 'purple' } : undefined} />;
    })}
  </div>;
}
const plugins = [defineDocxPlugin<DocxProposalSnapshot | null>({ id: 'host-proposals', createState: () => null,
  async onEvent(context, event) {
    if (event.type === 'proposal-change') probe.events.push(event);
    if (event.type !== 'load' && event.type !== 'proposal-change' && event.type !== 'layout-change') return;
    const snapshot = await probe.editor!.getProposals();
    context.setState(snapshot, snapshot.version);
  }, overlay: Overlay,
})];
const fonts = [{ family: 'Liberation Sans', src: fontUrl }];
const attachEditor = (editor: DocxEditorRef | null) => { probe.editor = editor; };
const measurementFontProvider = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  useEffect(() => { void pagedDocx(3, 12).then(setBuffer); }, []);
  return buffer && <DocxEditor ref={attachEditor} documentBuffer={buffer} legacySyncApi={arm === 'hybrid'}
    allowHostProposals showHostProposalsInSidebar commentsSidebarOpen plugins={plugins} fonts={fonts}
    measurementFontProvider={measurementFontProvider} showToolbar={false}
    onError={error => probe.errors.push(error.message)} onPluginError={error => probe.errors.push(String(error.error))} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
