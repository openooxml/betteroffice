import { rezipPartsToArrayBuffer, toBytes } from '../../docx/rezip/parts';

const OPENXML = 'http://schemas.openxmlformats.org';
const DRAWING = `${OPENXML}/drawingml/2006`;
const OFFICE = 'application/vnd.openxmlformats-officedocument';
const WORD = `${OFFICE}.wordprocessingml`;
const MS_WORD = 'http://schemas.microsoft.com/office/word';

export const ONE_SAVE_NS = {
  w: `${OPENXML}/wordprocessingml/2006/main`,
  r: `${OPENXML}/officeDocument/2006/relationships`,
  a: `${DRAWING}/main`,
  wp: `${DRAWING}/wordprocessingDrawing`,
  dgm: `${DRAWING}/diagram`,
  c: `${DRAWING}/chart`,
  mc: `${OPENXML}/markup-compatibility/2006`,
  w14: `${MS_WORD}/2010/wordml`,
  w15: `${MS_WORD}/2012/wordml`,
  w16cid: `${MS_WORD}/2016/wordml/cid`,
  w16se: `${MS_WORD}/2015/wordml/symex`,
  w16cex: `${MS_WORD}/2018/wordml/cex`,
  dsp: 'http://schemas.microsoft.com/office/drawing/2008/diagram',
  x: 'urn:one-save:probe',
  rels: `${OPENXML}/package/2006/relationships`,
  types: `${OPENXML}/package/2006/content-types`,
};

const NS = ONE_SAVE_NS;
const ns = (...prefixes: (keyof typeof NS)[]) =>
  prefixes.map((prefix) => `xmlns:${prefix}="${NS[prefix]}"`).join(' ');
const DATE = '2024-01-02T03:04:05Z';
const COMMENT_PARAGRAPH = '20000001';
const DURABLE_ID = '20000002';
const DOCUMENT = 'word/document.xml';
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

export const EDIT_PARAGRAPH_ID = '10000002';
export const DIAGRAM_RELS = {
  dm: { id: 'rIdData', type: 'diagramData', target: 'diagrams/data1.xml' },
  lo: { id: 'rIdLayout', type: 'diagramLayout', target: 'diagrams/layout1.xml' },
  qs: { id: 'rIdStyle', type: 'diagramQuickStyle', target: 'diagrams/quickStyle1.xml' },
  cs: { id: 'rIdColors', type: 'diagramColors', target: 'diagrams/colors1.xml' },
};
const EXTENDED_COMMENTS = {
  commentsExtended: 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
  commentsIds: 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
  commentsExtensible: 'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
};
export const EXTENDED_COMMENT_PARTS = Object.keys(EXTENDED_COMMENTS).map(
  (name) => `word/${name}.xml`
);

const POSITIONS = ['para', 'run', 'hyperlink'] as const;
const PREFIXES = ['w', 'mc', 'w14', 'w15', 'w16cid', 'w16se', 'x'] as const;
type ProbePosition = (typeof POSITIONS)[number];
const probe = (prefix: string, position: ProbePosition) => `P-${prefix}-${position}`;
export const ONE_SAVE_PROBES = [
  ...POSITIONS.flatMap((position) => PREFIXES.map((prefix) => probe(prefix, position))),
  'P-alternate', 'P-choice', 'P-fallback',
];

const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const para = (id: string, children: string) => `<w:p w14:paraId="${id}">${children}</w:p>`;
const alternate = (marker: string, choice: string, fallback: string) =>
  `<mc:AlternateContent x:val="${marker}"><mc:Choice Requires="w14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;

function probesAt(position: ProbePosition): string {
  const word = position === 'run'
    ? `<w:pgNum x:val="${probe('w', position)}"/>`
    : `<w:proofErr w:type="spellStart" x:val="${probe('w', position)}"/>${run('Word probe')}<w:proofErr w:type="spellEnd"/>`;
  const branch = position === 'run' ? '<w:t>Alternate</w:t>' : run('Alternate');
  return word + alternate(probe('mc', position), branch, branch) +
    PREFIXES.slice(2).map((prefix) => `<${prefix}:probeElem ${prefix}:val="${probe(prefix, position)}"/>`).join('');
}

function drawing(id: number, name: string, uri: string, graphic: string): string {
  return `<w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="${id}" name="${name}"/><a:graphic><a:graphicData uri="${uri}">${graphic}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

function move(kind: 'moveFrom' | 'moveTo', rangeId: number, revisionId: number): string {
  const text = kind === 'moveFrom' ? 'delText' : 't';
  const revision = `w:author="Author" w:date="${DATE}"`;
  return `<w:${kind}RangeStart w:id="${rangeId}" ${revision} w:name="Moved"/><w:${kind} w:id="${revisionId}" ${revision}><w:r><w:${text}>Moved text</w:${text}></w:r></w:${kind}><w:${kind}RangeEnd w:id="${rangeId}"/>`;
}

function notes(kind: 'footnote' | 'endnote', id: number, text: string): string {
  const paraId = kind === 'footnote' ? '30000001' : '30000002';
  return `<w:${kind}s ${ns('w', 'w14', 'mc')} mc:Ignorable="w14"><w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}><w:${kind} w:id="${id}">${para(paraId, `<w:r><w:${kind}Ref/></w:r>${run(text)}`)}</w:${kind}></w:${kind}s>`;
}

type Relationship = { id: string; type: string; target: string; external?: boolean };
const relationships = (entries: Relationship[]) =>
  `<Relationships xmlns="${NS.rels}">${entries.map(({ id, type, target, external }) =>
    `<Relationship Id="${id}" Type="${type.includes(':') ? type : `${NS.r}/${type}`}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`
  ).join('')}</Relationships>`;

export function oneSavePackage(): { bytes: Uint8Array; probes: readonly string[] } {
  const parts = new Map<string, Uint8Array>();
  const overrides: string[] = [];
  const set = (name: string, xml: string, type?: string) => {
    parts.set(name, toBytes(DECLARATION + xml));
    if (type) overrides.push(`<Override PartName="/${name}" ContentType="${type}+xml"/>`);
  };

  const diagramIds = Object.entries(DIAGRAM_RELS)
    .map(([attribute, { id }]) => `r:${attribute}="${id}"`).join(' ');
  const body = [
    para('10000001', run('Opening paragraph')),
    para(EDIT_PARAGRAPH_ID, run('EDIT ME')),
    para('10000003', drawing(1, 'Diagram', NS.dgm, `<dgm:relIds ${diagramIds}/>`)),
    para('10000004', move('moveFrom', 10, 11)),
    para('10000005', move('moveTo', 20, 21)),
    para('10000006', probesAt('para') + run('Paragraph probes')),
    para('10000007', `<w:r><w:t>Run probes</w:t>${probesAt('run')}${alternate(
      'P-alternate', '<w:t x:val="P-choice">Choice</w:t>', '<w:t x:val="P-fallback">Fallback</w:t>'
    )}</w:r>`),
    para('10000008', `<w:hyperlink r:id="rIdLink">${run('Link probes')}${probesAt('hyperlink')}</w:hyperlink>`),
    para('10000009', `<w:commentRangeStart w:id="1"/>${run('Commented text')}<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>`),
    para('1000000A', drawing(2, 'Chart', NS.c, '<c:chart r:id="rIdChart"/>')),
    para('1000000B', `${run('Notes')}<w:r><w:footnoteReference w:id="31"/></w:r><w:r><w:endnoteReference w:id="32"/></w:r>`),
    para('1000000C', `<w:bookmarkStart w:id="40" w:name="closing"/>${run('Closing paragraph')}<w:bookmarkEnd w:id="40"/>`),
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
  ].join('');
  set(DOCUMENT, `<w:document ${ns('w', 'r', 'a', 'wp', 'dgm', 'c', 'mc', 'w14', 'w15', 'w16cid', 'w16se', 'x')} mc:Ignorable="w14 w15 w16cid w16se x"><w:body>${body}</w:body></w:document>`, `${WORD}.document.main`);

  set('word/diagrams/data1.xml', `<dgm:dataModel ${ns('dgm', 'a', 'dsp')}><dgm:ptLst><dgm:pt modelId="1"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Diagram node</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/><dgm:extLst><a:ext uri="{C8258D5B-601F-41F4-9845-6A7AB8AC169B}"><dsp:dataModelExt relId="rIdDrawing" minVer="${NS.dgm}"/></a:ext></dgm:extLst></dgm:dataModel>`, `${OFFICE}.drawingml.diagramData`);
  set('word/diagrams/layout1.xml', `<dgm:layoutDef ${ns('dgm')} uniqueId="urn:one-save:layout" minVer="${NS.dgm}"><dgm:title val="Basic List"/><dgm:layoutNode name="diagram"><dgm:alg type="lin"/></dgm:layoutNode></dgm:layoutDef>`, `${OFFICE}.drawingml.diagramLayout`);
  set('word/diagrams/quickStyle1.xml', `<dgm:styleDef ${ns('dgm')} uniqueId="urn:one-save:style" minVer="${NS.dgm}"><dgm:title val="Simple style"/></dgm:styleDef>`, `${OFFICE}.drawingml.diagramStyle`);
  set('word/diagrams/colors1.xml', `<dgm:colorsDef ${ns('dgm', 'a')} uniqueId="urn:one-save:colors" minVer="${NS.dgm}"><dgm:title val="Simple colors"/></dgm:colorsDef>`, `${OFFICE}.drawingml.diagramColors`);
  set('word/diagrams/drawing1.xml', `<dsp:drawing ${ns('dsp', 'a')}><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name="Diagram"/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/><dsp:sp modelId="1"><dsp:nvSpPr><dsp:cNvPr id="1" name="Node"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Diagram node</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`, 'application/vnd.ms-office.drawingml.diagramDrawing');
  set('word/diagrams/_rels/data1.xml.rels', relationships([
    { id: 'rIdDrawing', type: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing', target: 'drawing1.xml' },
  ]));

  set('word/comments.xml', `<w:comments ${ns('w', 'w14', 'mc')} mc:Ignorable="w14"><w:comment w:id="1" w:author="Reviewer" w:date="${DATE}">${para(COMMENT_PARAGRAPH, run('A comment'))}</w:comment></w:comments>`, `${WORD}.comments`);
  set('word/commentsExtended.xml', `<w15:commentsEx ${ns('w15')}><w15:commentEx w15:paraId="${COMMENT_PARAGRAPH}" w15:done="0"/></w15:commentsEx>`, `${WORD}.commentsExtended`);
  set('word/commentsIds.xml', `<w16cid:commentsIds ${ns('w16cid')}><w16cid:commentId w16cid:paraId="${COMMENT_PARAGRAPH}" w16cid:durableId="${DURABLE_ID}"/></w16cid:commentsIds>`, `${WORD}.commentsIds`);
  set('word/commentsExtensible.xml', `<w16cex:commentsExtensible ${ns('w16cex')}><w16cex:commentExtensible w16cex:durableId="${DURABLE_ID}" w16cex:dateUtc="${DATE}"/></w16cex:commentsExtensible>`, `${WORD}.commentsExtensible`);

  set('word/charts/chart1.xml', `<c:chartSpace ${ns('c')}><c:chart><c:plotArea><c:pieChart><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>One</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>`, `${OFFICE}.drawingml.chart`);
  set('word/header1.xml', `<w:hdr ${ns('w', 'w14', 'mc')} mc:Ignorable="w14">${para('40000001', run('Header text'))}</w:hdr>`, `${WORD}.header`);
  set('word/footer1.xml', `<w:ftr ${ns('w', 'w14', 'mc')} mc:Ignorable="w14">${para('40000002', run('Footer text'))}</w:ftr>`, `${WORD}.footer`);
  set('word/footnotes.xml', notes('footnote', 31, 'Footnote text'), `${WORD}.footnotes`);
  set('word/endnotes.xml', notes('endnote', 32, 'Endnote text'), `${WORD}.endnotes`);
  set('docProps/core.xml', `<cp:coreProperties xmlns:cp="${OPENXML}/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Round-trip probes</dc:title><dcterms:modified xsi:type="dcterms:W3CDTF">${DATE}</dcterms:modified></cp:coreProperties>`, 'application/vnd.openxmlformats-package.core-properties');
  set('docProps/app.xml', `<Properties xmlns="${OPENXML}/officeDocument/2006/extended-properties"><Application>Host app</Application></Properties>`, `${OFFICE}.extended-properties`);

  set('_rels/.rels', relationships([
    { id: 'rIdDocument', type: 'officeDocument', target: DOCUMENT },
    { id: 'rIdCore', type: `${NS.rels}/metadata/core-properties`, target: 'docProps/core.xml' },
    { id: 'rIdApp', type: 'extended-properties', target: 'docProps/app.xml' },
  ]));
  set('word/_rels/document.xml.rels', relationships([
    ...Object.values(DIAGRAM_RELS),
    { id: 'rIdLink', type: 'hyperlink', target: 'https://example.invalid/', external: true },
    { id: 'rIdChart', type: 'chart', target: 'charts/chart1.xml' },
    { id: 'rIdHeader', type: 'header', target: 'header1.xml' },
    { id: 'rIdFooter', type: 'footer', target: 'footer1.xml' },
    { id: 'rIdFootnotes', type: 'footnotes', target: 'footnotes.xml' },
    { id: 'rIdEndnotes', type: 'endnotes', target: 'endnotes.xml' },
    { id: 'rIdComments', type: 'comments', target: 'comments.xml' },
    ...Object.entries(EXTENDED_COMMENTS).map(([name, type]) => ({ id: `rId${name}`, type, target: `${name}.xml` })),
  ]));
  set('[Content_Types].xml', `<Types xmlns="${NS.types}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>${overrides.join('')}</Types>`);

  return { bytes: new Uint8Array(rezipPartsToArrayBuffer(parts)), probes: ONE_SAVE_PROBES };
}
