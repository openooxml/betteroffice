import JSZip from 'jszip';
import { pptxFixture } from './format-fixtures';

export async function presentationFixture(options: { group?: boolean; runs?: number; fields?: boolean } = {}) {
  const zip = await JSZip.loadAsync(await pptxFixture());
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = 'http://schemas.openxmlformats.org/package/2006/relationships';
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<Relationships xmlns="${rels}"><Relationship Id="layout" Type="${r}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="notes" Type="${r}/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout xmlns:p="${p}" xmlns:a="${a}" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/notesSlides/notesSlide1.xml', `<p:notes xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Present the quarterly result.</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>`);
  let slide = await zip.file('ppt/slides/slide1.xml')!.async('string');
  slide = slide.replace('<p:nvPr/></p:nvSpPr>', '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>');
  if (options.runs) slide = slide.replace('Keep this paragraph.', '</a:t></a:r>' + Array.from({ length: options.runs }, (_, i) => `<a:r><a:rPr i="${i % 2}"/><a:t>${i}😀</a:t></a:r>`).join('') + '<a:r><a:t>tail');
  if (options.fields) slide = slide.replace('</a:p></p:txBody>', '<a:fld id="{AE409839-2DAE-4A29-9CE6-CAEC912BBF6D}" type="slidenum"><a:t>1</a:t></a:fld></a:p></p:txBody>');
  if (options.group) {
    const start = slide.indexOf('<p:sp>');
    const end = slide.indexOf('</p:sp>', start) + '</p:sp>'.length;
    const shape = slide.slice(start, end);
    slide = slide.slice(0, start) + `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/><a:chOff x="0" y="0"/><a:chExt cx="9144000" cy="6858000"/></a:xfrm></p:grpSpPr>${shape}</p:grpSp>` + slide.slice(end);
  }
  slide = slide.replace('</p:sld>', '<p:transition spd="slow"><p:fade/></p:transition></p:sld>');
  zip.file('ppt/slides/slide1.xml', slide);
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>'));
  return zip.generateAsync({ type: 'uint8array' });
}
