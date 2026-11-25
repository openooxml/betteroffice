import { expect, spyOn, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { createInProcessPair } from '../../../../shared/office-session/testing/inProcessTransport';
import { metafileWasmAssetUrl } from '../wasm/asset';
import { engineKind } from '../wasm/engine';
import { initWasm } from '../wasm/loader';
import { createPresentationCommandClient } from './commandClient';
import type { PresentationCommandClient } from './commandProtocol';
import { createPresentationSessionHost } from './host';

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

test.skipIf(engineKind === 'native')('command frames bound retries when the metafile WASM returns 404', async () => {
  await initWasm(await readFile(new URL('../wasm/generated/pptx_wasm_lean_bg.wasm', import.meta.url)));
  const bytes = await deckWithMetafile();
  const pair = createInProcessPair();
  createPresentationSessionHost(pair.host, { initWasm: async () => {} });

  const converterUrl = metafileWasmAssetUrl().href;
  const originalFetch = globalThis.fetch;
  let attempts = 0;
  const fetchSpy = spyOn(globalThis, 'fetch').mockImplementation(async (input, init) => {
    if (String(input instanceof Request ? input.url : input) !== converterUrl) return originalFetch(input, init);
    attempts++;
    return new Response(null, { status: 404 });
  });

  let client: PresentationCommandClient | undefined;
  try {
    client = await createPresentationCommandClient(bytes, {
      session: 'metafile-load-failure', incarnation: 'metafile-load-failure-worker', defaultFonts: false,
    }, pair.client);
    expect(client.state.frame).not.toBeNull();
    const slideId = client.state.slides[0]!.id;
    for (let request = 0; request < 50; request++) {
      expect((await client.thumbnail(slideId)).status).toBe('frame');
    }
    expect(attempts, 'Expected at least one converter WASM fetch').toBeGreaterThan(0);
    expect(attempts, `Converter WASM fetched ${attempts} times during open and 50 frame requests`).toBeLessThanOrEqual(3);
  } finally {
    try {
      await client?.dispose();
    } finally {
      pair.client.close();
      fetchSpy.mockRestore();
    }
  }
}, 10_000);
