import { expect, spyOn, test } from 'bun:test';
import { engines } from '#agents-engines';
import JSZip from 'jszip';
import { openDocx } from '../src';
import { renderDocxPage } from '../src/render';
import { fixture } from './fixture';

const w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const field = (name: string) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> ${name} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

async function fieldsAndNotes() {
  const zip = await JSZip.loadAsync(await fixture(120));
  const edit = async (part: string, change: (xml: string) => string) => zip.file(part, change(await zip.file(part)!.async('string')));
  zip.file('word/footer1.xml', `<w:ftr ${w}><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field('PAGE')}<w:r><w:t xml:space="preserve"> of </w:t></w:r>${field('NUMPAGES')}</w:p></w:ftr>`);
  zip.file('word/footnotes.xml', `<w:footnotes ${w}><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>A note.</w:t></w:r></w:p></w:footnote></w:footnotes>`);
  await edit('[Content_Types].xml', xml => xml.replace('</Types>', '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'));
  await edit('word/_rels/document.xml.rels', xml => xml.replace('</Relationships>', '<Relationship Id="ftr9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="fn9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>'));
  await edit('word/document.xml', xml => {
    const withNote = xml.replace('</w:p>', '<w:r><w:footnoteReference w:id="1"/></w:r></w:p>');
    const footer = '<w:footerReference w:type="default" r:id="ftr9"/>';
    if (withNote.includes('<w:sectPr')) return withNote.replace(/<w:sectPr([^>]*)>/, `<w:sectPr$1>${footer}`);
    return withNote.replace('</w:body>', `<w:sectPr>${footer}</w:sectPr></w:body>`);
  });
  return zip.generateAsync({ type: 'uint8array' });
}

async function laidOut(bytes: Uint8Array) {
  const { module } = await engines.load('docx');
  const session = await module.yrs.createYrsSession();
  session.openDocx(bytes, true);
  const document = session.materializeDocx()!;
  const source = module.layout.createRustMeasureSource({ engine: session, bundled: await module.layout.resolveDefaultFontProvider() });
  source.setCompat(document.package.settings?.compatibilityFlags);
  source.setEmbeddedFaces(await module.utils.extractEmbeddedFontFaces(document));
  const renderEnv = module.editor.documentRenderEnv(document);
  const requirements = JSON.parse(session.layoutFontRequirementsJson(JSON.stringify(module.editor.buildResidentRegionLayoutRequest(document, 0, renderEnv))));
  await source.prepareFontRequirements(requirements);
  const measurement = source.measurementConfigForRequirements(requirements)!;
  const { layout } = module.editor.computeLayout({ document, pageGap: 0, session, renderEnv, measurement });
  return { module, session, layout, measurement };
}

const painted = (page: unknown) => JSON.parse(JSON.stringify(page));

test('a page built alone matches the whole-document list, footer fields and notes included', async () => {
  const bytes = await fieldsAndNotes();
  const whole = await laidOut(bytes);
  const paged = await laidOut(bytes);
  try {
    const kernel = whole.module.editor.getLayoutKernelInputs(whole.layout)!;
    const expected = await whole.module.render.buildRustDisplayList({ ...kernel, layout: whole.layout, fontChains: whole.measurement.fontChains }, whole.session);
    expect(expected.pages.length).toBeGreaterThan(2);
    const fetchKernel = spyOn(paged.session, 'retainedKernelInputsJson');
    const display = paged.module.render.openPagedDisplay(paged.session, paged.measurement.fontChains);
    expect(display.pageCount).toBe(expected.pages.length);
    const order = [expected.pages.length - 1, 0, ...expected.pages.keys()];
    for (const index of order) expect(painted(display.page(index))).toEqual(painted(expected.pages[index]));
    expect(display.page(expected.pages.length)).toBeUndefined();
    expect(display.displayList.pages.filter(page => !page.unbuilt)).toHaveLength(1);
    expect(fetchKernel).not.toHaveBeenCalled();
    const footer = JSON.stringify(expected.pages[1]!.footer);
    expect(footer).toContain('NUMPAGES');
  } finally { whole.session.destroy(); paged.session.destroy(); }
});

test('the canvas path fetches no kernel inputs and leaves the caller session as it was', async () => {
  const { module } = await engines.load('docx');
  const { yrs } = module;
  let kernelFetches = 0;
  (module as { yrs: typeof yrs }).yrs = {
    ...yrs,
    createYrsSession: async (...args: Parameters<typeof yrs.createYrsSession>) => {
      const session = await yrs.createYrsSession(...args);
      const fetch = session.retainedKernelInputsJson;
      session.retainedKernelInputsJson = (revision) => { kernelFetches++; return fetch(revision); };
      return session;
    },
  };
  const doc = await openDocx(await fieldsAndNotes());
  const render = spyOn(doc.session, 'renderPng').mockRejectedValue(Object.assign(new Error('canvas unavailable'), { code: 'canvas-unavailable' }));
  const touched = (['setDisplayWindow', 'setDisplayRetainBuiltPages', 'buildDisplayListFrame', 'buildDisplayPagesFrame', 'layoutDocumentWithRegionsRetainedJson'] as const)
    .map(name => spyOn(doc.session, name));
  try {
    expect((await renderDocxPage(doc.session, 2)).pageCount).toBeGreaterThan(2);
    expect(kernelFetches).toBe(0);
    for (const spy of touched) expect(spy).not.toHaveBeenCalled();
    await expect(renderDocxPage(doc.session, 100_000)).rejects.toMatchObject({ code: 'PAGE_OUT_OF_RANGE' });
    for (const spy of touched) expect(spy).not.toHaveBeenCalled();
  } finally {
    (module as { yrs: typeof yrs }).yrs = yrs;
    for (const spy of touched) spy.mockRestore();
    render.mockRestore();
    doc.close();
  }
});
