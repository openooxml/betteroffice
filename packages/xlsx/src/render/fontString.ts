import type { TextCmd } from '../display-list/types';
import type { CellScriptFamilies } from '../../../../shared/office-session/xlsxScriptFonts';

const PX_PER_PT = 96 / 72;
const quoteFamily = (family: string) => `"${family.replace(/"/g, '\\"')}"`;

// css font shorthand: `[italic] [bold] <size>px <family>`. Workbook fonts are
// often not installed on the host (Calibri on macOS); without a fallback the
// browser's unknown-family default is serif.
export function fontString(cmd: TextCmd & { cssFontFamily?: string }, scriptFamilies?: CellScriptFamilies): string {
  const style = cmd.italic ? 'italic ' : '';
  const weight = cmd.bold ? 'bold ' : '';
  const family = cmd.chart ? cmd.cssFontFamily ?? (cmd.fontFamily || 'sans-serif') : cellFontFamily(cmd, scriptFamilies?.(cmd));
  return `${style}${weight}${fontSizePx(cmd)}px ${family}`;
}

function cellFontFamily(cmd: TextCmd & { cssFontFamily?: string }, scriptFamilies: string[] | undefined): string {
  const requested = cmd.fontFamily;
  const base = cmd.cssFontFamily ?? (requested ? `${quoteFamily(requested)}, sans-serif` : 'sans-serif');
  if (!scriptFamilies?.length) return base;
  const primary = requested?.trim().toLowerCase();
  let fallbacks = '';
  for (const name of scriptFamilies) {
    if (name.toLowerCase() !== primary) fallbacks += `${quoteFamily(name)}, `;
  }
  return base.replace(/sans-serif$/, `${fallbacks}sans-serif`);
}

export function fontSizePx(cmd: TextCmd): number {
  return cmd.fontSize * PX_PER_PT;
}
