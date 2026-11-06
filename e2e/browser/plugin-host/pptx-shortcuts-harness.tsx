import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import JSZip from 'jszip';
import {
  PptxEditor,
  definePptxPlugin,
  type PptxAsyncEditorApi,
  type PptxEditorApi,
  type PptxShortcuts,
  type ShortcutIssue,
} from '@betteroffice/pptx-react';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';

const legacy = new URLSearchParams(location.search).get('mode') === 'legacy';
const route = legacy ? { legacySyncApi: true, experimentalWorkerOpen: false } : {};
const fonts = [{ family: 'Liberation Sans', src: fontUrl }];

async function deck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  const shape = (id: number, text: string, y: number) => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="914400" y="${y}"/><a:ext cx="3657600" cy="914400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
    <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"><a:latin typeface="Liberation Sans"/></a:rPr>
    <a:t>${text}</a:t></a:r><a:endParaRPr sz="2400"/></a:p></p:txBody></p:sp>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${r}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/></p:sldIdLst>
    <p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', rels(`<Relationship Id="rId1" Type="${r}/slide" Target="slides/slide1.xml"/>`));
  zip.file('ppt/slides/slide1.xml', `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    ${shape(2, 'First box', 914400)}${shape(3, 'Target', 2743200)}
    </p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}

let open: (state: { file: Uint8Array; shortcuts: PptxShortcuts }) => void = () => {};
const probe = {
  editor: null as PptxEditorApi | PptxAsyncEditorApi | null,
  issues: [] as ShortcutIssue[],
  hostKeys: [] as { key: string; prevented: boolean }[],
  saves: 0,
  pluginRuns: 0,
  async load(shortcuts: PptxShortcuts) { open({ file: await deck(), shortcuts }); },
  state(id: 'bold' | 'italic') { return probe.editor?.commands.getState(id).active ?? null; },
  async texts() {
    const read = await probe.editor!.readContent();
    return (read as unknown as { stories: { text: string }[] }).stories.map((story) => story.text);
  },
};
(window as unknown as { __shortcuts: typeof probe }).__shortcuts = probe;
window.addEventListener('keydown', (event) => {
  if (!event.ctrlKey && !event.metaKey) return;
  probe.hostKeys.push({ key: event.key.toLowerCase(), prevented: event.defaultPrevented });
});
const plugin = definePptxPlugin({
  id: 'acme.keys',
  createState: () => null,
  commands: [{
    id: 'run', label: 'Run', mutatesDocument: false, shortcuts: ['Mod+Shift+K'],
    execute: () => { probe.pluginRuns += 1; return { ok: true, status: 'noop' }; },
  }],
} as never);
const plugins = [plugin] as never;

function Harness() {
  const [state, setState] = useState<{ file: Uint8Array; shortcuts: PptxShortcuts } | null>(null);
  open = setState;
  if (!state) return null;
  return <PptxEditor {...(route as object)} file={state.file} fonts={fonts} plugins={plugins}
    shortcuts={state.shortcuts} onShortcutIssue={(issue) => probe.issues.push(issue)}
    onSave={() => { probe.saves += 1; }} onReady={(api: PptxEditorApi | PptxAsyncEditorApi) => { probe.editor = api; }} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
