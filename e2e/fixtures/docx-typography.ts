export interface TypographyRow {
  id: string;
  font: string;
  family: string;
  text: string;
  ligatures?: string;
  contextualAlternates?: boolean;
  kerningMinPt?: number;
  rtl?: boolean;
}

export const typographyRows: TypographyRow[] = [
  { id: 'carlito-default', font: 'carlito', family: 'Calibri', text: 'fi ti' },
  { id: 'carlito-none', font: 'carlito', family: 'Calibri', text: 'fi ti', ligatures: 'none' },
  { id: 'carlito-joined', font: 'carlito', family: 'Calibri', text: 'fi ti', ligatures: 'standardContextual' },
  { id: 'carlito-av-default', font: 'carlito', family: 'Calibri', text: 'AV' },
  { id: 'carlito-av-kern', font: 'carlito', family: 'Calibri', text: 'AV', kerningMinPt: 8 },
  ...[
    { font: 'arabic', family: 'Noto Naskh Arabic', text: 'لا سلام', rtl: true },
    { font: 'devanagari', family: 'Noto Sans Devanagari', text: 'क्ष नमस्ते' },
  ].flatMap(face => ['none', 'all'].flatMap(ligatures => [false, true].map(contextualAlternates => ({
    ...face, id: `${face.font}-${ligatures}-${contextualAlternates}`, ligatures, contextualAlternates,
  })))),
];

export interface TypographySession {
  open_docx(bytes: Uint8Array, editable: boolean): unknown;
  decode_font(bytes: Uint8Array): Uint8Array;
  register_measure_font(bytes: Uint8Array): number;
  yrs_blocks_for_story(story: string, environment: string): string;
  measure_paragraph_json(input: string): string;
  build_display_list_json(input: string): string;
  free(): void;
}

export interface TypographyEngine {
  EditSession: new (client: number) => TypographySession;
  rezip_docx(parts: Record<string, Uint8Array>): Uint8Array;
  decodeFontFace(bytes: Uint8Array, codec: string): { font: Uint8Array };
}

export interface TypographyResult {
  id: string;
  measured: Array<[number, number]>;
  painted: Array<[number, number]>;
}

export async function inspectTypographyTable(input: {
  engine?: TypographyEngine;
  moduleUrl?: string;
  rows: TypographyRow[];
  fonts: Record<string, number[]>;
}): Promise<TypographyResult[]> {
  const imported = input.engine ? undefined : await import(input.moduleUrl!);
  if (imported) await imported.default();
  const engine: TypographyEngine = input.engine ?? imported;
  const encoder = new TextEncoder();
  const encode = (text: string) => encoder.encode(text);
  const unpack = (packed: string): Array<[number, number]> => {
    const bytes = Uint8Array.from(atob(packed), character => character.charCodeAt(0));
    const view = new DataView(bytes.buffer);
    let offset = 0;
    const varint = () => {
      let value = 0;
      let shift = 0;
      let byte: number;
      do {
        byte = view.getUint8(offset++);
        value |= (byte & 127) << shift;
        shift += 7;
      } while (byte & 128);
      return value;
    };
    const count = view.getUint32(offset, true);
    offset += 4;
    const glyphs: Array<[number, number]> = [];
    for (let run = 0; run < count; run++) {
      for (let field = 0; field < 4; field++) varint();
      offset++;
      const flags = view.getUint8(offset++);
      offset += 8;
      const count = varint();
      for (let index = 0; index < count; index++) {
        glyphs.push([view.getUint16(offset, true), view.getFloat32(offset + 2, true)]);
        offset += 6;
      }
      if (flags & 1) offset += count;
      if (flags & 2) {
        const bitmap = bytes.slice(offset, offset + Math.ceil(count / 8));
        offset += bitmap.length;
        for (let index = 0; index < count; index++) {
          if (bitmap[index >> 3] & (1 << (index & 7))) offset += 8;
        }
      }
    }
    if (offset !== bytes.length) throw new Error('Invalid packed glyph runs');
    return glyphs;
  };
  const results: TypographyResult[] = [];
  for (const row of input.rows) {
    const session = new engine.EditSession(713);
    try {
      const properties = [
        `<w:rFonts w:ascii="${row.family}" w:hAnsi="${row.family}" w:cs="${row.family}"/>`,
        '<w:sz w:val="24"/><w:szCs w:val="24"/>',
        row.kerningMinPt === undefined ? '' : `<w:kern w:val="${row.kerningMinPt * 2}"/>`,
        row.rtl ? '<w:rtl/>' : '',
        row.ligatures === undefined ? '' : `<w14:ligatures w14:val="${row.ligatures}"/>`,
        row.contextualAlternates === undefined ? '' : `<w14:cntxtAlts w14:val="${Number(row.contextualAlternates)}"/>`,
      ].join('');
      session.open_docx(engine.rezip_docx({
        '[Content_Types].xml': encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
        '_rels/.rels': encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),
        'word/document.xml': encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p>${row.rtl ? '<w:pPr><w:bidi/></w:pPr>' : ''}<w:r><w:rPr>${properties}</w:rPr><w:t>${row.text}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`),
      }), true);
      const fontIds = Object.fromEntries(Object.entries(input.fonts).map(([name, bytes]) => {
        const raw = new Uint8Array(bytes);
        const zstd = raw[0] === 0x28 && raw[1] === 0xb5 && raw[2] === 0x2f && raw[3] === 0xfd;
        const face = zstd ? engine.decodeFontFace(raw, 'zstd').font : raw;
        return [name, session.register_measure_font(session.decode_font(face))];
      }));
      const fontChains = { [`${row.family.toLowerCase()}|0|0`]: [fontIds[row.font]] };
      const blocks = JSON.parse(session.yrs_blocks_for_story('body', '{}'));
      const block = blocks.find((block: { kind: string }) => block.kind === 'paragraph');
      if (!block) throw new Error(`No paragraph for ${row.id}`);
      const measure = JSON.parse(session.measure_paragraph_json(JSON.stringify({
        block, maxWidth: 600, fontChains, authoritativeShaping: true,
        defaults: { fontSize: 12, fontFamily: row.family },
      })));
      const list = JSON.parse(session.build_display_list_json(JSON.stringify({
        measured: [{ block, measure }], fontChains, options: {},
        layout: { pages: [{
          number: 1, size: { w: 816, h: 1056 },
          margins: { top: 96, right: 96, bottom: 96, left: 96 },
          fragments: [{ kind: 'paragraph', blockId: block.id, x: 96, y: 96,
            width: 600, height: 100, fromLine: 0, toLine: measure.lines.length,
            pmStart: block.pmStart, pmEnd: block.pmEnd }],
        }] },
      })));
      const primitives = list.pages.flatMap((page: { primitives: unknown[] }) => page.primitives);
      if (primitives.some((run: { kind: string; text?: string }) => run.kind === 'text' && run.text)) {
        throw new Error(`Unexpected text fallback for ${row.id}`);
      }
      results.push({
        id: row.id,
        measured: measure.lines.flatMap((line: { glyphRuns: string }) => unpack(line.glyphRuns)),
        painted: primitives.filter((run: { kind: string }) => run.kind === 'glyphRun')
          .flatMap((run: { glyphs: Array<{ id: number; advance: number }> }) =>
            run.glyphs.map(glyph => [glyph.id, glyph.advance])),
      });
    } finally {
      session.free();
    }
  }
  return results;
}

export function assertTypographyTable(results: TypographyResult[]): void {
  for (const row of results) {
    if (!row.measured.length || row.measured.length !== row.painted.length) {
      throw new Error(`${row.id}: measured and painted glyph counts differ`);
    }
    row.measured.forEach(([id, advance], index) => {
      const painted = row.painted[index];
      if (id !== painted[0] || Math.abs(advance - painted[1]) > 0.002) {
        throw new Error(`${row.id}: measured and painted glyph ${index} differ`);
      }
    });
  }
  const find = (id: string) => {
    const row = results.find(row => row.id === id);
    if (!row) throw new Error(`Missing ${id}`);
    return row;
  };
  if (find('carlito-default').measured.length !== 5
    || JSON.stringify(find('carlito-none').measured) !== JSON.stringify(find('carlito-default').measured)
    || find('carlito-joined').measured.length >= 5) {
    throw new Error('Calibri optional ligature settings were not applied');
  }
  const advance = (id: string) => find(id).measured.reduce((sum, glyph) => sum + glyph[1], 0);
  if (advance('carlito-av-kern') >= advance('carlito-av-default')) {
    throw new Error('Calibri kerning threshold was not applied');
  }
  for (const face of ['arabic', 'devanagari']) {
    const rows = results.filter(row => row.id.startsWith(`${face}-`));
    const signature = (row: TypographyResult) => JSON.stringify([row.measured, row.painted]);
    if (rows.length !== 4 || rows.some(row => signature(row) !== signature(rows[0]))) {
      throw new Error(`${face}: optional typography changed complex-script shaping`);
    }
  }
}
