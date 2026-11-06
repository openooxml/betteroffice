import { createCanvas, GlobalFonts } from '@napi-rs/canvas';
import { createFontProvider, type BundledFontSource } from '@betteroffice/fonts';
import { decodeFontFace, paintDisplayList, type DisplayList, type Viewport, type WorkbookHandle } from '@betteroffice/xlsx';

export async function renderXlsxRange(
  handle: WorkbookHandle,
  sheet: number,
  range: string,
  viewport: Viewport,
  scale: number,
  provider: BundledFontSource = createFontProvider(undefined, { decoder: decodeFontFace }),
): Promise<Uint8Array> {
  const active = handle.sheetInfo().activeSheet;
  let list: DisplayList;
  let fonts: ReturnType<WorkbookHandle['renderPngFonts']>;
  try {
    if (sheet !== active) handle.setActiveSheet(sheet);
    list = handle.displayList(viewport);
    fonts = handle.renderPngFonts({ sheet, range }, { scale });
  } finally {
    if (sheet !== active) handle.setActiveSheet(active);
  }
  const requests = new Map(fonts.map(font => [JSON.stringify(font), font]));
  for (const cmd of list.commands) {
    if (cmd.op !== 'text') continue;
    const family = (cmd.chart ? cmd.fontFamily?.split(',')[0]?.trim().replace(/^(['"])(.*)\1$/, '$2') : cmd.fontFamily) || 'sans-serif';
    const font = { family, bold: cmd.bold === true, italic: cmd.italic === true };
    requests.set(JSON.stringify(font), font);
  }
  for (const font of requests.values()) {
    let bytes: ArrayBuffer | undefined;
    try { bytes = await provider.resolve(font.family, font.bold, font.italic)?.(); } catch {}
    bytes ??= await provider.resolveLastResort(font.family, font.bold, font.italic)();
    GlobalFonts.register(Buffer.from(bytes), font.family);
  }
  const canvas = createCanvas(Math.ceil(viewport.width * scale), Math.ceil(viewport.height * scale));
  const ctx = canvas.getContext('2d');
  paintDisplayList(ctx as unknown as CanvasRenderingContext2D, list, scale);
  ctx.globalCompositeOperation = 'destination-over';
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  return new Uint8Array(await canvas.encode('png'));
}
