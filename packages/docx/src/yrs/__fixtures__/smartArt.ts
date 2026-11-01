const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';

export function smartArtFixture(anchored: boolean, options: { bodyNamespace?: boolean; extensionSize?: number; hidden?: boolean } = {}) {
  const container = anchored
    ? '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="19" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>91440</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>45720</wp:posOffset></wp:positionV>'
    : '<wp:inline distT="0" distB="0" distL="0" distR="0">';
  const drawing = `<w:drawing xmlns:a="${A}"${options.bodyNamespace ? '' : ` xmlns:dgm="${DGM}"`}>${container}<wp:extent cx="5486400" cy="3200400"/>` +
    '<wp:effectExtent l="19050" t="9525" r="0" b="38100"/>' +
    (anchored ? '<wp:wrapSquare wrapText="bothSides"/>' : '') +
    '<wp:docPr id="71" name="Diagram &amp; process" descr="Keep diagram metadata">' +
    `<a:extLst><a:ext uri="urn:synthetic:docPr"><keep:metadata xmlns:keep="urn:synthetic:keep" value="yes"/></a:ext></a:extLst></wp:docPr>` +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    `<a:graphic><a:graphicData uri="${DGM}"><dgm:relIds r:dm="rId10" r:lo="rId11" r:qs="rId12" r:cs="rId13"/>` +
    `<a:extLst><a:ext uri="urn:synthetic:diagram"><keep:payload xmlns:keep="urn:synthetic:keep">${'x'.repeat(options.extensionSize ?? 0)}Retain me</keep:payload></a:ext></a:extLst>` +
    `</a:graphicData></a:graphic></wp:${anchored ? 'anchor' : 'inline'}></w:drawing>`;
  const drawingRun = '<w:r w:rsidR="0099AABB"><w:rPr><w:noProof/>' + (options.hidden ? '<w:vanish/>' : '') + '<w:lang w:val="en-US"/></w:rPr>' + drawing + '</w:r>';
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:a="${A}" xmlns:wp="${WP}" ${options.bodyNamespace ? '' : `xmlns:dgm="${DGM}" `}xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body${options.bodyNamespace ? ` xmlns:dgm="${DGM}"` : ''}>` +
    '<w:p w14:paraId="10000001" w:rsidR="00A1B2C3"><w:r w:rsidR="00A1B2C3"><w:t xml:space="preserve">Process overview before </w:t></w:r>' +
    drawingRun + '<w:r><w:t xml:space="preserve"> and after the diagram.</w:t></w:r></w:p>' +
    '<w:p w14:paraId="10000002"><w:r><w:t>Second paragraph stays.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>';
  const diagramParts = [
    ['data1.xml', 'diagramData', `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst>` +
      '<dgm:pt modelId="root" type="doc"><dgm:prSet loTypeId="urn:microsoft.com/office/officeart/2005/8/layout/process1"/></dgm:pt>' +
      '<dgm:pt modelId="item"><dgm:prSet/><dgm:t><a:bodyPr/><a:p><a:r><a:t>Plan</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst>' +
      '<dgm:cxnLst><dgm:cxn modelId="edge" srcId="root" destId="item" srcOrd="0" destOrd="0"/></dgm:cxnLst>' +
      `<dgm:extLst><a:ext uri="${DSP}"><dsp:dataModelExt xmlns:dsp="${DSP}" relId="rId14"/></a:ext></dgm:extLst></dgm:dataModel>`],
    ['layout1.xml', 'diagramLayout', `<dgm:layoutDef xmlns:dgm="${DGM}" uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/process1"><dgm:layoutNode name="process"/></dgm:layoutDef>`],
    ['quickStyle1.xml', 'diagramStyle', `<dgm:styleDef xmlns:dgm="${DGM}" uniqueId="synthetic-style"><dgm:title val="Keep style"/></dgm:styleDef>`],
    ['colors1.xml', 'diagramColors', `<dgm:colorsDef xmlns:dgm="${DGM}" uniqueId="synthetic-colors"><dgm:title val="Keep colors"/></dgm:colorsDef>`],
    ['drawing1.xml', 'diagramDrawing', `<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>` +
      '<dsp:sp modelId="item"><dsp:nvSpPr><dsp:cNvPr id="2" name="Plan"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1600000" cy="960000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></dsp:spPr>' +
      '<dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Plan</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>'],
  ];
  const diagramRelationships = [
    ['rId10', 'diagramData', 'data1.xml'], ['rId11', 'diagramLayout', 'layout1.xml'],
    ['rId12', 'diagramQuickStyle', 'quickStyle1.xml'], ['rId13', 'diagramColors', 'colors1.xml'],
    ['rId14', 'diagramDrawing', 'drawing1.xml'],
  ];
  const parts = new Map<string, string>([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      diagramParts.map(([name, type]) => `<Override PartName="/word/diagrams/${name}" ContentType="${type === 'diagramDrawing' ? 'application/vnd.ms-office.drawingml.diagramDrawing+xml' : `application/vnd.openxmlformats-officedocument.drawingml.${type}+xml`}"/>`).join('') + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="${RELS}"><Relationship Id="main" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/document.xml', document],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="${RELS}">` + diagramRelationships.map(([id, type, name]) =>
      `<Relationship Id="${id}" Type="${type === 'diagramDrawing' ? 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing' : `${R}/${type}`}" Target="diagrams/${name}"/>`).join('') + '</Relationships>'],
    ['word/diagrams/_rels/data1.xml.rels', `<Relationships xmlns="${RELS}"><Relationship Id="rId14" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="drawing1.xml"/></Relationships>`],
    ...diagramParts.map(([name, , xml]): [string, string] => [`word/diagrams/${name}`, xml!]),
  ]);
  return { parts, drawingRun, drawing };
}

export function smartArtDrawingRun(xml: string): string | undefined {
  return xml.match(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g)?.find(run => run.includes('<w:drawing'));
}

export function smartArtTargets(document: string, relationships: string): string[] {
  const ids = document.match(/<dgm:relIds\b[^>]*\/>/)?.[0];
  if (!ids) return [];
  return ['dm', 'lo', 'qs', 'cs'].flatMap(name => {
    const id = ids.match(new RegExp(`r:${name}="([^"]+)"`))?.[1];
    const target = relationships.match(new RegExp(`<Relationship\\b[^>]*Id="${id}"[^>]*Target="([^"]+)"`))?.[1];
    return target ? [`word/${target}`] : [];
  });
}

export function smartArtNamespaces(xml: string): Array<string | undefined> {
  const scopes: Array<Map<string, string>> = [new Map()];
  const found: Array<string | undefined> = [];
  for (const token of xml.matchAll(/<\/?[\w:]+\b(?:"[^"]*"|'[^']*'|[^'">])*>/g)) {
    const tag = token[0];
    if (tag.startsWith('</')) { scopes.pop(); continue; }
    const scope = new Map(scopes[scopes.length - 1]);
    for (const attr of tag.matchAll(/\sxmlns(?::([\w.-]+))?\s*=\s*(["'])(.*?)\2/g)) scope.set(attr[1] ?? '', attr[3]!);
    if (tag.startsWith('<dgm:relIds')) found.push(scope.get('dgm'));
    if (!tag.endsWith('/>')) scopes.push(scope);
  }
  return found;
}
