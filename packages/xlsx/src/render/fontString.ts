import type { TextCmd } from '../display-list/types';
import { cellScriptFaces } from './cellScriptFonts';

const PX_PER_PT = 96 / 72;
const quoteFamily = (family: string) => `"${family.replace(/"/g, '\\"')}"`;

// css font shorthand: `[italic] [bold] <size>px <family>`. Workbook fonts are
// often not installed on the host (Calibri on macOS); without a fallback the
// browser's unknown-family default is serif.
export function fontString(cmd: TextCmd & { cssFontFamily?: string }): string {
  const style = cmd.italic ? 'italic ' : '';
  const weight = cmd.bold ? 'bold ' : '';
  const family = cmd.cssFontFamily ?? (cmd.chart ? cmd.fontFamily || 'sans-serif' : cellFontFamily(cmd));
  return `${style}${weight}${fontSizePx(cmd)}px ${family}`;
}

function cellFontFamily(cmd: TextCmd): string {
  const requested = cmd.fontFamily;
  const primary = requested?.trim().toLowerCase();
  let families = requested ? `${quoteFamily(requested)}, ` : '';
  for (const { face } of cellScriptFaces(cmd)) {
    if (face.family.toLowerCase() !== primary) families += `${quoteFamily(face.family)}, `;
  }
  return `${families}sans-serif`;
}

export function fontSizePx(cmd: TextCmd): number {
  return cmd.fontSize * PX_PER_PT;
}
