import JSZip from 'jszip';

export async function fontPreflightFixture(): Promise<Uint8Array> {
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (entries: [string, string, string][]) =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.map(([id, kind, target]) => `<Relationship Id="${id}" Type="${r}/${kind}" Target="${target}"/>`).join('')}</Relationships>`;
  const zip = new JSZip();
  const parts: Record<string, string> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + [
      ['/ppt/presentation.xml', 'presentation.main'], ['/ppt/slides/one.xml', 'slide'],
      ['/ppt/slideLayouts/one.xml', 'slideLayout'], ['/ppt/slideMasters/one.xml', 'slideMaster'],
    ].map(([part, kind]) => `<Override PartName="${part}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.${kind}+xml"/>`).join('') + '<Override PartName="/ppt/theme/one.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
    '_rels/.rels': rels([['main', 'officeDocument', 'ppt/presentation.xml']]),
    'ppt/presentation.xml': `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="master"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': rels([['slide', 'slide', 'slides/one.xml'], ['master', 'slideMaster', 'slideMasters/one.xml']]),
    'ppt/slides/one.xml': `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"><a:latin typeface="+mn-lt"/></a:rPr><a:t>漢かな</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
    'ppt/slides/_rels/one.xml.rels': rels([['layout', 'slideLayout', '../slideLayouts/one.xml']]),
    'ppt/slideLayouts/one.xml': `<p:sldLayout xmlns:p="${p}"><p:cSld><p:spTree/></p:cSld></p:sldLayout>`,
    'ppt/slideLayouts/_rels/one.xml.rels': rels([['master', 'slideMaster', '../slideMasters/one.xml']]),
    'ppt/slideMasters/one.xml': `<p:sldMaster xmlns:p="${p}" xmlns:r="${r}"><p:cSld><p:spTree/></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="layout"/></p:sldLayoutIdLst></p:sldMaster>`,
    'ppt/slideMasters/_rels/one.xml.rels': rels([['theme', 'theme', '../theme/one.xml'], ['layout', 'slideLayout', '../slideLayouts/one.xml']]),
    'ppt/theme/one.xml': `<a:theme xmlns:a="${a}"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Heading"/></a:majorFont><a:minorFont><a:latin typeface="Body"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`,
  };
  for (const [name, xml] of Object.entries(parts)) zip.file(name, xml);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
