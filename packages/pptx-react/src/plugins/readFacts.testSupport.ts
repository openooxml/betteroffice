import JSZip from 'jszip';

export async function readFactsDeck(): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${[
    ['presentation', 'presentation.main'], ['slides/slide1', 'slide'], ['slideLayouts/slideLayout1', 'slideLayout'], ['slideMasters/slideMaster1', 'slideMaster'],
  ].map(([path, type]) => `<Override PartName="/ppt/${path}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.${type}+xml"/>`).join('')}</Types>`);
  const relationships = (type: string, target: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/${type}" Target="${target}"/></Relationships>`;
  zip.file('_rels/.rels', relationships('officeDocument', 'ppt/presentation.xml'));
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId2"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', relationships('slide', 'slides/slide1.xml').replace('</Relationships>', `<Relationship Id="rId2" Type="${rel}/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>`));
  zip.file('ppt/slides/_rels/slide1.xml.rels', relationships('slideLayout', '../slideLayouts/slideLayout1.xml'));
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', relationships('slideMaster', '../slideMasters/slideMaster1.xml'));
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', relationships('slideLayout', '../slideLayouts/slideLayout1.xml'));
  const group = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
  zip.file('ppt/slideMasters/slideMaster1.xml', `<p:sldMaster ${ns}><p:cSld><p:spTree>${group}</p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:hf sldNum="0" ftr="0" dt="0"/></p:sldMaster>`);
  const footer = (id: number, kind: string, runs: string) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${kind}"/><p:cNvSpPr/><p:nvPr><p:ph type="${kind}"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${runs}</a:p></p:txBody></p:sp>`;
  const footerShapes = footer(3, 'ftr', '<a:r><a:t>Footer</a:t></a:r>') + footer(4, 'sldNum', '<a:r><a:t>Seite </a:t></a:r><a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:t>9</a:t></a:fld>');
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout ${ns} type="title"><p:cSld name="Heading layout"><p:spTree>${group}</p:spTree></p:cSld><p:hf hdr="0"/></p:sldLayout>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns} show="0"><p:cSld><p:spTree>${group}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Heading"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="2400000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${['Heading', 'Next'].map((text) => `<a:p><a:pPr><a:buAutoNum type="alphaLcParenR"/></a:pPr><a:r><a:rPr sz="2400"/><a:t>${text}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp>${footerShapes}</p:spTree></p:cSld></p:sld>`);
  return zip.generateAsync({ type: 'uint8array' });
}
