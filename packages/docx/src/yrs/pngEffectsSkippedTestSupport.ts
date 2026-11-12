/**
 * The kinds of effect each PNG engine reports for a page, one docx per case.
 * Both engines' PNG tests read this table, so a kind one engine reports is
 * named the same way by the other. Each case reports the same kinds whether
 * layout emits its runs as text or as glyph runs.
 */

import type { PngSkippedEffectKind } from '../../../../shared/office-session/pngExportError';

export type PngEngineKind = 'wasm' | 'native';

export interface EffectsSkippedDiagnostic {
  code: 'effects-skipped';
  page: number;
  effects: PngSkippedEffectKind[];
}

export interface EffectFixture {
  name: string;
  parts: Record<string, Uint8Array>;
  /** The kinds page 0 reports, per engine; an empty list means no diagnostic. */
  skipped: Record<PngEngineKind, PngSkippedEffectKind[]>;
}

export const ONE_PIXEL_PNG = new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYPj/HwADAgH/5ncLrgAAAABJRU5ErkJggg==', 'base64'));

const encoder = new TextEncoder();

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
].join(' ');

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const HEADER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';

/** Bytes no image decoder accepts: a PNG signature cut off inside its first chunk. */
export const UNDECODABLE_PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13]);

interface DocxSource {
  body: string;
  header?: string;
  /** The bytes of `media/image1.png`, or null to leave the part out; a one-pixel PNG by default. */
  image?: Uint8Array | null;
  /** The bytes of `media/image2.png`, related as `rIdImage2`; no such part by default. */
  secondImage?: Uint8Array;
}

/** A one-section docx whose body (and optional default header) is the given markup. */
export function effectDocx({ body, header, image = ONE_PIXEL_PNG, secondImage }: DocxSource): Record<string, Uint8Array> {
  const headerOverride = header
    ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    : '';
  const headerRelationship = header ? `<Relationship Id="rIdHeader" Type="${HEADER_REL}" Target="header1.xml"/>` : '';
  const sectPr = header ? '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>' : '<w:sectPr/>';
  const secondRelationship = secondImage ? `<Relationship Id="rIdImage2" Type="${IMAGE_REL}" Target="media/image2.png"/>` : '';
  const parts: Record<string, Uint8Array> = {
    '[Content_Types].xml': encoder.encode(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${headerOverride}</Types>`),
    '_rels/.rels': encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
    'word/_rels/document.xml.rels': encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${IMAGE_REL}" Target="media/image1.png"/>${secondRelationship}${headerRelationship}</Relationships>`),
    'word/document.xml': encoder.encode(`<w:document ${NAMESPACES}><w:body>${body}${sectPr}</w:body></w:document>`),
  };
  if (image) parts['word/media/image1.png'] = image;
  if (secondImage) parts['word/media/image2.png'] = secondImage;
  if (header) {
    parts['word/header1.xml'] = encoder.encode(`<w:hdr ${NAMESPACES}>${header}</w:hdr>`);
    parts['word/_rels/header1.xml.rels'] = encoder.encode(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdWatermark" Type="${IMAGE_REL}" Target="media/image1.png"/></Relationships>`);
  }
  return parts;
}

const styledRun = (properties: string, text = 'Effect') => `<w:p><w:r><w:rPr>${properties}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const SOLID_FILL = '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>';
const pictureFill = (relationship: string) => `<a:blipFill><a:blip r:embed="${relationship}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
const PICTURE_FILL = pictureFill('rIdImage');

const shapeWithEffects = (effects: string, fill = SOLID_FILL) => `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapNone/><wp:docPr id="1" name="Shape"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill}<a:effectLst>${effects}</a:effectLst></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;

const pictureWithBlip = (blipChildren: string) => `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="2" name="Picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage">${blipChildren}</a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

const washedOutWatermark = '<w:p><w:r><w:pict><v:shape id="WatermarkPicture1" o:spid="_x0000_s2049" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:96pt;height:48pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f"><v:imagedata r:id="rIdWatermark" o:title="" gain="19661f" blacklevel="22938f"/></v:shape></w:pict></w:r></w:p>';

/** A body holding one inline picture of the one-pixel PNG, with no effects. */
export const PICTURE_BODY = pictureWithBlip('');

const GLOW = '<w14:glow w14:rad="63500"><w14:srgbClr w14:val="FF0000"/></w14:glow>';
const SHADOW_THEN_SOFT_EDGE = '<a:outerShdw blurRad="38100" dist="25400" dir="2700000"><a:srgbClr val="000000"/></a:outerShdw><a:softEdge rad="12700"/>';
const TWO_UNDRAWN_SHAPE_EFFECTS = '<a:reflection blurRad="0" stA="50000" endA="0" endPos="35000" dist="0" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/><a:softEdge rad="12700"/>';

export const PLAIN_FIXTURE: EffectFixture = {
  name: 'plain text',
  parts: effectDocx({ body: '<w:p><w:r><w:t>Plain</w:t></w:r></w:p>' }),
  skipped: { wasm: [], native: [] },
};

export const EFFECT_FIXTURES: EffectFixture[] = [
  {
    name: 'a classic outline run',
    parts: effectDocx({ body: styledRun('<w:outline/>') }),
    skipped: { wasm: [], native: ['text.outline'] },
  },
  {
    name: 'an emphasis-mark run',
    parts: effectDocx({ body: styledRun('<w:em w:val="dot"/>') }),
    skipped: { wasm: [], native: ['text.emphasisMark'] },
  },
  {
    name: 'an animated text effect run',
    parts: effectDocx({ body: styledRun('<w:effect w:val="blinkBackground"/>') }),
    skipped: { wasm: ['text.effect'], native: ['text.effect'] },
  },
  {
    name: 'a glowing run',
    parts: effectDocx({ body: styledRun(GLOW) }),
    skipped: { wasm: [], native: ['text.modernEffects'] },
  },
  {
    name: 'a run with a classic shadow, which both engines draw',
    parts: effectDocx({ body: styledRun('<w:shadow/>') }),
    skipped: { wasm: [], native: [] },
  },
  {
    name: 'a shadowed glowing run, whose glow costs the classic shadow on the canvas',
    parts: effectDocx({ body: styledRun(`<w:shadow/>${GLOW}`) }),
    skipped: { wasm: ['text.shadow'], native: ['text.modernEffects'] },
  },
  {
    name: 'a dotted tab leader, which layout draws as a dotted decoration',
    parts: effectDocx({ body: '<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="3600"/></w:tabs></w:pPr><w:r><w:t>Chapter</w:t></w:r><w:r><w:tab/><w:t>9</w:t></w:r></w:p>' }),
    skipped: { wasm: [], native: [] },
  },
  {
    name: 'a shape with a soft edge after its shadow',
    parts: effectDocx({ body: shapeWithEffects(SHADOW_THEN_SOFT_EDGE) }),
    skipped: { wasm: ['shape.effects'], native: ['shape.effects'] },
  },
  {
    name: 'a shape with two effects neither engine draws',
    parts: effectDocx({ body: shapeWithEffects(TWO_UNDRAWN_SHAPE_EFFECTS) }),
    skipped: { wasm: ['shape.effects'], native: ['shape.effects'] },
  },
  {
    name: 'an outlined, emphasised, animated run beside a shape that loses an effect',
    parts: effectDocx({ body: `${styledRun('<w:outline/><w:em w:val="dot"/><w:effect w:val="blinkBackground"/>')}${shapeWithEffects(SHADOW_THEN_SOFT_EDGE)}` }),
    skipped: {
      wasm: ['text.effect', 'shape.effects'],
      native: ['text.outline', 'text.emphasisMark', 'text.effect', 'shape.effects'],
    },
  },
  {
    name: 'a washed-out VML picture in the header, whose washout does not reach the display list',
    parts: effectDocx({ body: '<w:p><w:r><w:t>Plain</w:t></w:r></w:p>', header: washedOutWatermark }),
    skipped: { wasm: [], native: [] },
  },
  {
    name: 'a duotone picture, whose effect never reaches the page',
    parts: effectDocx({ body: pictureWithBlip('<a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FFFFFF"/></a:duotone>') }),
    skipped: { wasm: [], native: [] },
  },
];

/** Page 0 is plain; page 1 holds one animated text effect run. */
export const MIXED_FIXTURE = {
  name: 'one plain page then one page with an effect',
  parts: effectDocx({ body: '<w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:rPr><w:effect w:val="blinkBackground"/></w:rPr><w:t>Effect</w:t></w:r></w:p>' }),
  effectPage: 1,
  skipped: ['text.effect'] as PngSkippedEffectKind[],
};

/** Shapes whose picture fill is not drawn, each also carrying effects neither engine draws. */
export const UNDRAWN_PICTURE_FILL_FIXTURES = [
  {
    name: 'a picture-filled shape whose picture part is missing',
    parts: effectDocx({ body: shapeWithEffects(SHADOW_THEN_SOFT_EDGE, PICTURE_FILL), image: null }),
  },
  {
    name: 'a picture-filled shape whose picture cannot be decoded',
    parts: effectDocx({ body: shapeWithEffects(SHADOW_THEN_SOFT_EDGE, PICTURE_FILL), image: UNDECODABLE_PNG }),
  },
];

const OUTER_SHADOW = '<a:outerShdw blurRad="38100" dist="25400" dir="2700000"><a:srgbClr val="000000"/></a:outerShdw>';

/**
 * An animated run beside two shapes filled with pictures of one type: the
 * first, with a soft edge neither engine draws, cannot be decoded; the second
 * can, and carries an outer shadow and the given further effects.
 */
const twoPictures = (further: string) => effectDocx({
  body: `${styledRun('<w:effect w:val="blinkBackground"/>')}${shapeWithEffects('<a:softEdge rad="12700"/>', pictureFill('rIdImage'))}${shapeWithEffects(`${OUTER_SHADOW}${further}`, pictureFill('rIdImage2'))}`,
  image: UNDECODABLE_PNG,
  secondImage: ONE_PIXEL_PNG,
});

/** Pages holding two pictures of one type, one unavailable, and the kinds each reports on either engine. */
export const TWO_PICTURE_FIXTURES = [
  {
    name: 'the unavailable picture reports nothing beside one whose effects are drawn',
    parts: twoPictures(''),
    skipped: ['text.effect'] as PngSkippedEffectKind[],
  },
  {
    name: 'the unavailable picture reports nothing beside one that loses an effect',
    parts: twoPictures('<a:softEdge rad="12700"/>'),
    skipped: ['text.effect', 'shape.effects'] as PngSkippedEffectKind[],
  },
];

/** The diagnostics a page with these kinds fires: none, or exactly one, its kinds in the listed order. */
export function expectedDiagnostics(page: number, skipped: PngSkippedEffectKind[]): EffectsSkippedDiagnostic[] {
  return skipped.length ? [{ code: 'effects-skipped', page, effects: skipped }] : [];
}

/** The `effects-skipped` diagnostics among those reported, as reported. */
export function effectsSkipped(diagnostics: ReadonlyArray<{ code: string }>): EffectsSkippedDiagnostic[] {
  return diagnostics.filter((diagnostic): diagnostic is EffectsSkippedDiagnostic => diagnostic.code === 'effects-skipped');
}
