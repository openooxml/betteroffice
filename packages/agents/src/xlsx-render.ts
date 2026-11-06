import { engines, loadCanvas } from '#agents-engines';
import { createFontProvider, type BundledFontSource } from '@betteroffice/fonts';
import type { DisplayList, Viewport, WorkbookHandle } from '@betteroffice/xlsx';

export async function renderXlsxRange(
  handle: WorkbookHandle,
  sheet: number,
  range: string,
  viewport: Viewport,
  scale: number,
  provider?: BundledFontSource,
): Promise<Uint8Array> {
  const { createCanvas, GlobalFonts } = await loadCanvas();
  const { module } = await engines.load('xlsx');
  provider ??= createFontProvider(undefined, { decoder: module.decodeFontFace });
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
  const familyOf = (cmd: { chart?: boolean; fontFamily?: string }) => (cmd.chart ? cmd.fontFamily?.split(',')[0]?.trim().replace(/^(['"])(.*)\1$/, '$2') : cmd.fontFamily) || 'sans-serif';
  for (const cmd of list.commands) {
    if (cmd.op !== 'text') continue;
    const font = { family: familyOf(cmd), bold: cmd.bold === true, italic: cmd.italic === true };
    requests.set(JSON.stringify(font), font);
  }
  const render = ++renders;
  const aliases = new Map<string, string>();
  const keys: NonNullable<ReturnType<typeof GlobalFonts.register>>[] = [];
  try {
    for (const font of requests.values()) {
      let bytes: ArrayBuffer | undefined;
      try { bytes = await provider.resolve(font.family, font.bold, font.italic)?.(); } catch {}
      bytes ??= await provider.resolveLastResort(font.family, font.bold, font.italic, 'word')();
      let alias = aliases.get(font.family);
      if (!alias) aliases.set(font.family, alias = `BetterOffice Preview ${render} ${aliases.size}`);
      const key = GlobalFonts.register(Buffer.from(bytes), alias);
      if (key) keys.push(key);
    }
    const painted = { ...list, commands: list.commands.map(cmd => {
      if (cmd.op !== 'text') return cmd;
      const alias = aliases.get(familyOf(cmd));
      if (!alias) return cmd;
      return { ...cmd, fontFamily: cmd.chart ? [`"${alias}"`, ...(cmd.fontFamily?.split(',').slice(1) ?? [])].join(',') : alias };
    }) };
    const canvas = createCanvas(Math.ceil(Math.fround(Math.fround(viewport.width) * Math.fround(scale))), Math.ceil(Math.fround(Math.fround(viewport.height) * Math.fround(scale))));
    const ctx = canvas.getContext('2d');
    module.paintDisplayList(ctx as unknown as CanvasRenderingContext2D, painted, scale);
    ctx.globalCompositeOperation = 'destination-over';
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    return new Uint8Array(await canvas.encode('png'));
  } finally {
    for (const key of keys) GlobalFonts.remove(key);
  }
}

let renders = 0;

/** The size the engine gives a range: its edges and the scaled extent in 32-bit floats, as the native and wasm engines compute them. */
export function rangeExtent(first: { x: number; y: number }, last: { x: number; y: number; w: number; h: number }, scale: number) {
  const f = Math.fround;
  const width = f(f(last.x + last.w) - f(first.x));
  const height = f(f(last.y + last.h) - f(first.y));
  return { width, height, pixelWidth: Math.ceil(f(width * f(scale))), pixelHeight: Math.ceil(f(height * f(scale))) };
}
