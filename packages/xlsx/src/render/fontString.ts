import type { TextCmd } from '../display-list/types';

const PX_PER_PT = 96 / 72;

// css font shorthand: `[italic] [bold] <size>px <family>`. Workbook fonts are
// often not installed on the host (Calibri on macOS); without a fallback the
// browser's unknown-family default is serif.
export function fontString(cmd: TextCmd & { cssFontFamily?: string }): string {
  const style = cmd.italic ? 'italic ' : '';
  const weight = cmd.bold ? 'bold ' : '';
  const family = cmd.cssFontFamily ?? (cmd.fontFamily
    ? cmd.chart
      ? cmd.fontFamily
      : `"${cmd.fontFamily.replace(/"/g, '\\"')}", sans-serif`
    : 'sans-serif');
  return `${style}${weight}${fontSizePx(cmd)}px ${family}`;
}

export function fontSizePx(cmd: TextCmd): number {
  return cmd.fontSize * PX_PER_PT;
}
