import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { metafileWasmAssetUrl } from '../wasm/asset';
import { engineKind } from '../wasm/engine';
import { initWasm } from '../wasm/loader';
import { openPresentationSession, type PresentationSession } from './client';
import { createPresentationSessionHost } from './host';
import { commandMetafileSvgModule } from './metafileSvg';

async function deckWithMetafile(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="emf" ContentType="image/x-emf"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  </Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
    <p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
  </p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:pic><p:nvPicPr><p:cNvPr id="2" name="Metafile"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    </p:pic>
  </p:spTree></p:cSld></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<Relationships xmlns="${rels}"><Relationship Id="picture" Type="${rel}/image" Target="../media/picture.emf"/></Relationships>`);
  zip.file('ppt/media/picture.emf', await readFile(new URL('../../../../crates/pptx-render/tests/fixtures/metafile-text.emf', import.meta.url)));
  return zip.generateAsync({ type: 'uint8array' });
}

function sessionWorker(): Worker {
  const pair = createInProcessPair();
  createPresentationSessionHost(pair.host, { initWasm: async () => {} });
  const target = new EventTarget();
  pair.client.listen((data) => target.dispatchEvent(new MessageEvent('message', { data })));
  return Object.assign(target, {
    postMessage: (message: unknown, transfer?: Transferable[]) => pair.client.post(message, transfer),
    terminate: () => pair.client.close(),
  }) as unknown as Worker;
}

test.skipIf(engineKind === 'native')('the first session frame resolves with a pending metafile WASM fetch', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_lean_bg.wasm', import.meta.url)));
  const bytes = await deckWithMetafile();
  const converterUrl = metafileWasmAssetUrl().href;
  const originalFetch = globalThis.fetch;
  const pending = new Promise<Response>(() => {});
  let requested = false;
  const fetchSpy = spyOn(globalThis, 'fetch').mockImplementation((input, init) => {
    if (String(input instanceof Request ? input.url : input) !== converterUrl) return originalFetch(input, init);
    requested = true;
    return pending;
  });

  let client: PresentationSession | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    const firstFrame = openPresentationSession(bytes, { worker: sessionWorker }).then(async (session) => {
      client = session;
      await session.call.frame(0);
      return 'resolved' as const;
    });
    const timeout = new Promise<'hung'>((resolve) => {
      timer = setTimeout(() => resolve('hung'), 5_000);
    });
    const outcome = await Promise.race([firstFrame, timeout]);
    const phase = client ? 'First frame' : 'Presentation open';
    const message = outcome === 'hung'
      ? `${phase} hung for 5 s (converter fetch started: ${requested})`
      : 'Presentation open and first frame resolved within 5 s';
    expect(outcome, message).toBe('resolved');
    expect(requested, 'Expected the first frame to request the converter WASM').toBe(true);
  } finally {
    clearTimeout(timer);
    try {
      await client?.dispose();
    } finally {
      fetchSpy.mockRestore();
    }
  }
}, 10_000);

test.skipIf(engineKind === 'native')('a session frame carries SVG after the converter fetch outlasts the first frame cap', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_lean_bg.wasm', import.meta.url)));
  const bytes = await deckWithMetafile();
  const converterUrl = metafileWasmAssetUrl();
  const converterWasm = new Uint8Array(await readFile(converterUrl));
  const originalFetch = globalThis.fetch;
  let attempts = 0;
  const fetchSpy = spyOn(globalThis, 'fetch').mockImplementation(async (input, init) => {
    if (String(input instanceof Request ? input.url : input) !== converterUrl.href) return originalFetch(input, init);
    attempts++;
    await new Promise((resolve) => setTimeout(resolve, 3_500));
    return new Response(converterWasm, { headers: { 'Content-Type': 'application/wasm' } });
  });
  const load = spyOn(commandMetafileSvgModule, 'load');

  let client: PresentationSession | undefined;
  try {
    client = await openPresentationSession(bytes, { worker: sessionWorker });
    const assetId = 'ppt/media/picture.emf';
    const first = await client.call.frame(0);
    expect(first.media.get(assetId)).toHaveLength(0);
    await load.mock.results[0]!.value;
    const later = await client.call.frame(0);
    expect(new TextDecoder().decode(later.media.get(assetId)?.subarray(0, 256))).toMatch(/^\s*<svg[\s>]/);
    expect(attempts).toBe(1);
  } finally {
    try {
      await client?.dispose();
    } finally {
      load.mockRestore();
      fetchSpy.mockRestore();
    }
  }
}, 10_000);
