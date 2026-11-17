import JSZip from 'jszip';

/** Faces supplied by the tests, as paths from the repository root; each maps A, B and C to a distinct shape. */
export const PNG_FONT_FACES = {
  regular: 'crates/betteroffice-fonts/tests/fixtures/Synthetic-Regular.ttf',
  bold: 'crates/betteroffice-fonts/tests/fixtures/Synthetic-Bold.ttf',
} as const;

/** Bytes no engine can paint with. */
export const PNG_BROKEN_FACE = [1, 2, 3];

export type PngFontFaceName = keyof typeof PNG_FONT_FACES | 'broken';

export const PNG_FONT_TARGET = { sheet: 0, viewport: { x: 0, y: 0, width: 300, height: 150 } };

/** Each case is one `renderPng` call, made in this order on one workbook. */
export const PNG_FONT_CASES: { name: string; fonts?: { family: string; face: PngFontFaceName }[] }[] = [
  { name: 'absent' },
  { name: 'empty', fonts: [] },
  { name: 'absent after empty' },
  { name: 'regular', fonts: [{ family: 'Calibri', face: 'regular' }] },
  { name: 'bold', fonts: [{ family: 'Calibri', face: 'bold' }] },
  { name: 'absent after supplied faces' },
  { name: 'uncovered', fonts: [{ family: 'Not On The Sheet', face: 'broken' }] },
  { name: 'broken', fonts: [{ family: 'Calibri', face: 'broken' }] },
  { name: 'absent again' },
];

/** The behaviour every engine must show for the cases above. */
export const PNG_FONT_OUTCOMES: ({ rule: string } & (
  | { differs: [string, string] }
  | { equals: [string, string] }
  | { rejects: { case: string; code: string; family: string } }
))[] = [
  { rule: 'a supplied face paints the sheet font', differs: ['regular', 'absent'] },
  { rule: 'the supplied face decides what is painted', differs: ['regular', 'bold'] },
  { rule: 'an explicit empty list replaces the bundled faces', differs: ['empty', 'absent'] },
  { rule: 'a family the list does not cover is painted as with an empty list', equals: ['uncovered', 'empty'] },
  { rule: 'an empty list applies to its own export only', equals: ['absent after empty', 'absent'] },
  { rule: 'supplied faces apply to their own export only', equals: ['absent after supplied faces', 'absent'] },
  { rule: 'a rejected export leaves later exports unchanged', equals: ['absent again', 'absent'] },
  { rule: 'a supplied face that cannot paint rejects naming the family', rejects: { case: 'broken', code: 'font-unavailable', family: 'Calibri' } },
];

export type PngFontResult = { png: string } | { code?: string; message: string };

/** The rules of `PNG_FONT_OUTCOMES` the results break, one line each; empty when every rule holds. */
export function pngFontOutcomeFailures(results: Record<string, PngFontResult>): string[] {
  const failures: string[] = [];
  const png = (name: string, rule: string) => {
    const result = results[name];
    if (result && 'png' in result) return result.png;
    failures.push(`${rule}: "${name}" did not render (${result ? `${result.code ?? ''} ${result.message}` : 'missing'})`);
    return undefined;
  };
  for (const outcome of PNG_FONT_OUTCOMES) {
    if ('rejects' in outcome) {
      const result = results[outcome.rejects.case];
      if (!result || 'png' in result) failures.push(`${outcome.rule}: "${outcome.rejects.case}" rendered`);
      else if (result.code !== outcome.rejects.code || !result.message.includes(`"${outcome.rejects.family}"`)) {
        failures.push(`${outcome.rule}: "${outcome.rejects.case}" rejected with ${result.code ?? 'no code'}: ${result.message}`);
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

/** One sheet: A1 holds Latin text and A2 CJK and Devanagari text, all in Calibri. */
export async function pngFontsWorkbook(): Promise<Uint8Array> {
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.file('xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rels}"><sheets><sheet name="Fonts" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rels}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rels}/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="1"><font><sz val="24"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs></styleSheet>`);
  zip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><cols><col min="1" max="1" width="40" customWidth="1"/></cols><sheetData><row r="1" ht="40" customHeight="1"><c r="A1" t="inlineStr"><is><t>ABCABC</t></is></c></row><row r="2" ht="40" customHeight="1"><c r="A2" t="inlineStr"><is><t>日本語 नमस्ते</t></is></c></row></sheetData></worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}
