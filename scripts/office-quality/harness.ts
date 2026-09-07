import JSZip from 'jszip';
import { createFontProvider } from '../../packages/fonts/src/cdn';

const api = window as any;
const format = new URLSearchParams(location.search).get('format');
let capture: (index: number) => Promise<string>;
let printCapture: any;

async function fontsFor(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const families = new Set(['Arial', 'Calibri']);
  for (const entry of Object.values(zip.files)) {
    if (!/^(?:ppt\/.*|xl\/styles)\.xml$/.test(entry.name)) continue;
    const xml = new DOMParser().parseFromString(await entry.async('string'), 'text/xml');
    for (const node of xml.querySelectorAll('latin, name')) {
      const family = node.getAttribute('typeface') ?? node.getAttribute('val');
      if (family && !family.startsWith('+')) families.add(family);
    }
  }
  if (families.size > 32) throw new Error('Too many font families for this capture');
  const provider = createFontProvider();
  const faces = [];
  for (const family of families) {
    for (const [bold, italic] of [
      [false, false],
      [true, false],
      [false, true],
      [true, true],
    ]) {
      const load =
        provider.resolve(family, bold, italic) ??
        provider.resolveLastResort(family, bold, italic);
      const buffer = await load();
      const face = new FontFace(family, buffer.slice(0), {
        weight: bold ? '700' : '400',
        style: italic ? 'italic' : 'normal',
      });
      document.fonts.add(await face.load());
      faces.push({ family, bold, italic, bytes: new Uint8Array(buffer) });
    }
  }
  return faces;
}

async function xlsxPrintMetrics(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  const xml = async (path: string) =>
    new DOMParser().parseFromString((await zip.file(path)?.async('string')) ?? '<root/>', 'text/xml');
  const styles = await xml('xl/styles.xml');
  const normal = styles.querySelector('cellStyleXfs > xf');
  const font =
    styles.querySelectorAll('fonts > font')[Number(normal?.getAttribute('fontId') ?? 0)];
  const family = font?.querySelector('name')?.getAttribute('val') ?? 'Calibri';
  const size = Number(font?.querySelector('sz')?.getAttribute('val') ?? 11);
  const dpi = 72;
  const context = document.createElement('canvas').getContext('2d')!;
  context.font = `${(size * dpi) / 72}px "${family.replace(/"/g, '\\"')}"`;
  const bounds = context.measureText('0123456789');
  const maxDigitWidth = Math.max(
    ...[...'0123456789'].map((digit) => Math.round(context.measureText(digit).width))
  );
  const workbook = await xml('xl/workbook.xml');
  const relationships = await xml('xl/_rels/workbook.xml.rels');
  return Promise.all(
    [...workbook.querySelectorAll('sheet')].map(async (sheet) => {
      const id = sheet.getAttribute('r:id');
      const target = [...relationships.querySelectorAll('Relationship')]
        .find((rel) => rel.getAttribute('Id') === id)
        ?.getAttribute('Target');
      if (!target) throw new Error('Missing worksheet relationship');
      const path = new URL(
        target,
        'https://package.invalid/xl/workbook.xml'
      ).pathname.slice(1);
      const format = (await xml(path)).querySelector('sheetFormatPr');
      return {
        dpi,
        maxDigitWidth,
        fontSizePt: size,
        fontFamily: family,
        fontAscent: bounds.fontBoundingBoxAscent,
        fontDescent: bounds.fontBoundingBoxDescent,
        defaultRowHeightPt: Number(
          format?.getAttribute('defaultRowHeight') ??
            ((bounds.fontBoundingBoxAscent + bounds.fontBoundingBoxDescent + 1) * 72) /
              dpi
        ),
        ...(format?.hasAttribute('defaultColWidth')
          ? {
              defaultColumnWidth: Number(format.getAttribute('defaultColWidth')),
            }
          : {}),
      };
    })
  );
}

function cell(value: string) {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/.exec(value);
  if (!match) throw new Error('Invalid print range');
  const col = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const row = Number(match[2]) - 1;
  if (col >= 16384 || row >= 1048576)
    throw new Error('Print range exceeds worksheet limits');
  return { row, col };
}

api.oracleInit = async (input: number[], useFonts: boolean, profile: any) => {
  const bytes = new Uint8Array(input);
  const fonts = useFonts ? await fontsFor(bytes) : [];
  let pages: number;
  if (format === 'pptx') {
    const { initWasm, openPresentation, paintSlide, sizeCanvasForSlide } = await import(
      '@betteroffice/pptx'
    );
    await initWasm();
    const handle = openPresentation(bytes, { fonts });
    pages = handle.snapshot().slides.length;
    capture = async (index) => {
      const frame = handle.layoutSlide(index);
      const canvas = document.createElement('canvas');
      sizeCanvasForSlide(canvas, frame, 150 / 96, 1);
      const images = new Map<string, ImageBitmap>();
      try {
        await paintSlide(canvas.getContext('2d')!, frame, 150 / 96, 1, {
          resolveImage: async (path: string) => {
            if (!images.has(path))
              images.set(
                path,
                await createImageBitmap(new Blob([handle.mediaBytes(path).slice()]))
              );
            return images.get(path)!;
          },
        });
        return canvas.toDataURL('image/png');
      } finally {
        for (const bitmap of images.values()) bitmap.close();
      }
    };
  } else if (format === 'xlsx') {
    const { initWasm, openWorkbook, paintDisplayList } = await import(
      '@betteroffice/xlsx'
    );
    await initWasm();
    const handle = openWorkbook(bytes);
    const printMetrics = await xlsxPrintMetrics(bytes);
    printCapture = {
      mode:
        typeof handle.printDisplayList === 'function' ? 'print-range' : 'screen-range',
      metrics: printMetrics,
    };
    if (
      !profile?.pages?.length ||
      !Number.isFinite(profile.scale_percent) ||
      profile.scale_percent < 10 ||
      profile.scale_percent > 100 ||
      !Number.isFinite(profile.margin_pt) ||
      profile.margin_pt < 0
    )
      throw new Error('A recorded XLSX print profile is required');
    pages = profile.pages.length;
    capture = async (index) => {
      const page = profile.pages[index];
      if (
        !Number.isInteger(page.sheet) ||
        page.sheet < 0 ||
        page.sheet >= handle.sheetInfo().sheetNames.length
      )
        throw new Error('Invalid worksheet index');
      handle.setActiveSheet(page.sheet);
      const info = handle.sheetInfo();
      if ((info.frozenRows || info.frozenCols) && typeof handle.printDisplayList !== 'function')
        throw new Error('XLSX print capture does not support frozen panes');
      const parts = page.range.split(':');
      if (parts.length !== 2) throw new Error('A rectangular print range is required');
      const start = cell(parts[0]);
      const end = cell(parts[1]);
      if (end.row < start.row || end.col < start.col)
        throw new Error('Reversed print range');
      const printable =
        typeof handle.printDisplayList === 'function'
          ? handle.printDisplayList(
              page.sheet,
              page.range,
              printMetrics[page.sheet],
              profile.gridlines !== false
            )
          : undefined;
      const from = handle.cellPosition(page.sheet, start.row, start.col);
      const to = handle.cellPosition(page.sheet, end.row + 1, end.col + 1);
      const width = printable?.width ?? to.x - from.x;
      const height = printable?.height ?? to.y - from.y;
      const scale = ((150 / 96) * profile.scale_percent) / 100;
      const margin = (profile.margin_pt * 150) / 72;
      if (
        ![page.width_px, page.height_px].every(
          (n) => Number.isInteger(n) && n > 0 && n <= 5000
        ) ||
        width <= 0 ||
        height <= 0 ||
        width * scale > page.width_px - margin * 2 + 1 ||
        height * scale > page.height_px - margin * 2 + 1
      )
        throw new Error('Print range does not fit the recorded page');
      const canvas = document.createElement('canvas');
      canvas.width = page.width_px;
      canvas.height = page.height_px;
      const context = canvas.getContext('2d')!;
      context.fillStyle = '#fff';
      context.fillRect(0, 0, canvas.width, canvas.height);
      if (printable) {
        context.save();
        context.beginPath();
        context.rect(margin, margin, width * scale, height * scale);
        context.clip();
        paintDisplayList(context, printable, scale, { x: margin, y: margin });
        context.restore();
      } else {
        const content = document.createElement('canvas');
        content.width = Math.ceil(width * scale);
        content.height = Math.ceil(height * scale);
        paintDisplayList(
          content.getContext('2d')!,
          handle.displayList({ x: from.x, y: from.y, width, height }),
          scale
        );
        context.drawImage(content, margin, margin);
      }
      return canvas.toDataURL('image/png');
    };
  } else throw new Error('Unsupported capture format');
  if (!Number.isInteger(pages) || pages < 1 || pages > 100)
    throw new Error('Invalid page count');
  return {
    pages,
    errors: [],
    fontLoads: fonts.map(({ family, bold, italic }) => ({
      family,
      bold,
      italic,
      ok: true,
    })),
    capture_profile: profile,
    ...(printCapture ? { print_capture: printCapture } : {}),
  };
};
api.oraclePage = (index: number) => capture(index);
api.oracleReady = true;
