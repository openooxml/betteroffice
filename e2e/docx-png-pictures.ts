import JSZip from 'jszip';

export type PicturePart = 'none' | 'drawn' | 'missing' | 'corrupt';

/** Each case is one `renderPng` of page 0 of a document holding one picture of the given kind. */
export const PICTURE_CASES: { name: string; part: PicturePart; strictPictures?: boolean }[] = [
  { name: 'no picture', part: 'none' },
  { name: 'drawn', part: 'drawn' },
  { name: 'drawn strict', part: 'drawn', strictPictures: true },
  { name: 'missing', part: 'missing' },
  { name: 'missing strict', part: 'missing', strictPictures: true },
  { name: 'corrupt', part: 'corrupt' },
  { name: 'corrupt strict', part: 'corrupt', strictPictures: true },
];

/** A picture as the diagnostics name it. */
export interface ReportedPicture { id: string; reason: string }

/** One `onDiagnostic` call, as the runner received it. */
export interface ReportedDiagnostic { code: string; pictures: ReportedPicture[] }

/** One export: the PNG's digest, or the rejection; with every `onDiagnostic` call it made. */
export type PictureResult = ({ png: string } | { code?: string; message: string }) & { diagnostics: ReportedDiagnostic[] };

/** The behaviour every engine must show for the cases above. */
export const PICTURE_OUTCOMES: ({ rule: string } & (
  | { differs: [string, string] }
  | { equals: [string, string] }
  | { reports: { case: string; reasons: string[] } }
  | { rejects: { case: string; code: string; reason: string; named: string } }
))[] = [
  { rule: 'a picture that draws is on the page', differs: ['drawn', 'no picture'] },
  { rule: 'a picture that draws is not reported', reports: { case: 'drawn', reasons: [] } },
  { rule: 'a missing picture leaves a placeholder on the page', differs: ['missing', 'no picture'] },
  { rule: 'a missing picture is reported as missing', reports: { case: 'missing', reasons: ['missing'] } },
  { rule: 'a picture the decoder refuses leaves a placeholder on the page', differs: ['corrupt', 'no picture'] },
  { rule: 'a picture the decoder refuses is reported as decode-failed', reports: { case: 'corrupt', reasons: ['decode-failed'] } },
  { rule: 'strictPictures leaves a page whose pictures draw as it is', equals: ['drawn strict', 'drawn'] },
  { rule: 'strictPictures rejects a missing picture, naming it, with no PNG and no diagnostic', rejects: { case: 'missing strict', code: 'picture-unavailable', reason: 'missing', named: 'missing' } },
  { rule: 'strictPictures rejects a refused picture, naming it, with no PNG and no diagnostic', rejects: { case: 'corrupt strict', code: 'picture-unavailable', reason: 'decode-failed', named: 'corrupt' } },
];

/** The rules of `PICTURE_OUTCOMES` the results break, one line each; empty when every rule holds. */
export function pictureOutcomeFailures(results: Record<string, PictureResult>): string[] {
  const failures: string[] = [];
  const png = (name: string, rule: string) => {
    const result = results[name];
    if (result && 'png' in result) return result.png;
    failures.push(`${rule}: "${name}" did not render (${result && 'message' in result ? `${result.code ?? ''} ${result.message}` : 'missing'})`);
    return undefined;
  };
  for (const outcome of PICTURE_OUTCOMES) {
    if ('rejects' in outcome) {
      const { case: name, code, reason, named } = outcome.rejects;
      const result = results[name];
      const id = reportedIds(results[named]);
      if (!result || 'png' in result) failures.push(`${outcome.rule}: "${name}" rendered`);
      else if (result.code !== code || id.length !== 1 || !result.message.includes(`${id[0]} (${reason})`)) {
        failures.push(`${outcome.rule}: "${name}" rejected with ${result.code ?? 'no code'}: ${result.message}, where "${named}" reported ${JSON.stringify(id)}`);
      } else if (result.diagnostics.length) failures.push(`${outcome.rule}: "${name}" also reported ${JSON.stringify(result.diagnostics)}`);
      continue;
    }
    if ('reports' in outcome) {
      const { case: name, reasons } = outcome.reports;
      if (png(name, outcome.rule) === undefined) continue;
      const diagnostics = results[name]!.diagnostics;
      const reported = diagnostics.flatMap((diagnostic) => diagnostic.pictures);
      const expected = reasons.length ? 1 : 0;
      if (diagnostics.length !== expected || diagnostics.some((diagnostic) => diagnostic.code !== 'picture-unavailable')
        || JSON.stringify(reported.map((picture) => picture.reason).sort()) !== JSON.stringify([...reasons].sort())
        || reported.some((picture) => !picture.id)) {
        failures.push(`${outcome.rule}: "${name}" reported ${JSON.stringify(diagnostics)}`);
      }
      continue;
    }
    const [a, b] = 'differs' in outcome ? outcome.differs : outcome.equals;
    const first = png(a, outcome.rule);
    const second = png(b, outcome.rule);
    if (first === undefined || second === undefined) continue;
    if ('differs' in outcome && first === second) failures.push(`${outcome.rule}: "${a}" and "${b}" are identical`);
    if ('equals' in outcome && first !== second) failures.push(`${outcome.rule}: "${a}" and "${b}" differ`);
  }
  return failures;
}

/** The ids a lenient export reported, in order. */
function reportedIds(result: PictureResult | undefined): string[] {
  return result ? result.diagnostics.flatMap((diagnostic) => diagnostic.pictures.map((picture) => picture.id)) : [];
}

/** A 4 by 4 opaque red PNG. */
const RED_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAEElEQVR42mP4z8AARwzEcQCukw/xOF6MEQAAAABJRU5ErkJggg==';

/** One page with a line of text and, unless `part` is `none`, one inline picture that `part` describes. */
export async function pictureDocument(part: PicturePart): Promise<Uint8Array> {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const picture = part === 'none' ? '' : '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1828800" cy="914400"/><wp:docPr id="1" name="Picture"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${part === 'none' ? '' : `<Relationship Id="rIdImage" Type="${rel}/image" Target="media/image.png"/>`}</Relationships>`);
  zip.file('word/document.xml', `<w:document xmlns:w="${w}" xmlns:r="${rel}"><w:body><w:p><w:r><w:t>Pictures</w:t></w:r></w:p>${picture}</w:body></w:document>`);
  if (part === 'drawn') zip.file('word/media/image.png', RED_PNG, { base64: true });
  if (part === 'corrupt') zip.file('word/media/image.png', new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13]));
  return zip.generateAsync({ type: 'uint8array' });
}
