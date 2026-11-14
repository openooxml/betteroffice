import { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { PptxEditor } from '@betteroffice/pptx-react';
import { slicedFacesSupported } from '../../../packages/pptx/src/session/commandSlicedFonts';

const params = new URLSearchParams(location.search);
const sliced = params.get('sliced') !== '0';

const LINES: [string, string, string][] = [
  ['ja-JP', 'MS Gothic', 'ひらがな、カタカナと漢字の混ざった文章です。'],
  ['zh-CN', 'SimSun', '这是一段简体中文的示例文字。'],
  ['zh-TW', 'PMingLiU', '這是一段繁體中文的範例文字。'],
  ['ko-KR', 'Malgun Gothic', '한국어 예시 문장입니다. 한글과 漢字.'],
];

async function deck() {
  const zip = new JSZip();
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  const run = ([lang, face, text]: [string, string, string], bold: boolean) =>
    `<a:p><a:r><a:rPr lang="${lang}" altLang="en-US" sz="2800"${bold ? ' b="1"' : ''}><a:ea typeface="${face}"/></a:rPr><a:t>${text}</a:t></a:r></a:p>`;
  const slides = LINES.map((line, index) => {
    const order = [line, ...LINES.filter((other) => other !== line)];
    return `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="Text ${index + 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="8229600" cy="4114800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${order.map((entry, at) => run(entry, at === 0)).join('')}</p:txBody>
      </p:sp></p:spTree></p:cSld></p:sld>`;
  });
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
    ${slides.map((_, index) => `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}
  </Types>`);
  zip.file('_rels/.rels', rels(`<Relationship Id="rId1" Type="${r}/officeDocument" Target="ppt/presentation.xml"/>`));
  zip.file('ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}">
    <p:sldIdLst>${slides.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 1}"/>`).join('')}</p:sldIdLst>
    <p:sldSz cx="9144000" cy="5143500"/><p:notesSz cx="6858000" cy="9144000"/>
  </p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', rels(slides.map((_, index) =>
    `<Relationship Id="rId${index + 1}" Type="${r}/slide" Target="slides/slide${index + 1}.xml"/>`).join('')));
  slides.forEach((slide, index) => zip.file(`ppt/slides/slide${index + 1}.xml`, slide));
  return zip.generateAsync({ type: 'uint8array' });
}

document.getElementById('verdict')!.textContent = sliced && slicedFacesSupported() ? 'sliced' : 'whole faces';
document.getElementById('option')!.textContent = sliced ? 'on' : 'off';
const toggle = document.getElementById('toggle') as HTMLAnchorElement;
toggle.textContent = sliced ? 'turn the option off' : 'turn the option on';
toggle.href = `?sliced=${sliced ? 0 : 1}`;

function Check() {
  const [file, setFile] = useState<{ key: number; bytes: Uint8Array } | null>(null);
  useEffect(() => {
    void deck().then((bytes) => setFile({ key: 0, bytes }));
    const input = document.getElementById('file') as HTMLInputElement;
    input.onchange = () => {
      const chosen = input.files?.[0];
      if (chosen) void chosen.arrayBuffer().then((bytes) => setFile({ key: Date.now(), bytes: new Uint8Array(bytes) }));
    };
  }, []);
  if (!file) return null;
  return <PptxEditor key={file.key} file={file.bytes} fonts={[]} experimentalSlicedFonts={sliced} />;
}

createRoot(document.getElementById('sliced-fonts-root')!).render(<Check />);
