import JSZip from 'jszip';

export async function readFactsDeck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${[
    ['presentation', 'presentation.main'], ['slides/slide1', 'slide'], ['slideLayouts/slideLayout1', 'slideLayout'],
  ].map(([path, type]) => `<Override PartName="/ppt/${path}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.${type}+xml"/>`).join('')}</Types>`);
  const relationships = (type: string, target: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/${type}" Target="${target}"/></Relationships>`;
  zip.file('_rels/.rels', relationships('officeDocument', 'ppt/presentation.xml'));
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', relationships('slide', 'slides/slide1.xml'));
  zip.file('ppt/slides/_rels/slide1.xml.rels', relationships('slideLayout', '../slideLayouts/slideLayout1.xml'));
  const group = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout ${ns} type="title"><p:cSld name="Heading layout"><p:spTree>${group}</p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns} show="0"><p:cSld><p:spTree>${group}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Heading"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="2400000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${['Heading', 'Next'].map((text) => `<a:p><a:pPr><a:buAutoNum type="alphaLcParenR"/></a:pPr><a:r><a:rPr sz="2400"/><a:t>${text}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}
