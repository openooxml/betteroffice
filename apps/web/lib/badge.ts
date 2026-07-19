import { LOGOS } from "./logos";

const CHAR_WIDTHS: Record<string, number> = {
  " ": 3.5, "/": 3.6, ".": 3.5, ",": 3.5, ":": 3.5, "-": 4,
  i: 3.2, j: 3.2, l: 3.2, f: 3.6, t: 3.6, r: 4.2,
  m: 9.8, w: 8.6, M: 9.6, W: 10.2,
};

const LOGO_SIZE = 14;
const LOGO_GAP = 4;

function textWidth(text: string): number {
  let width = 0;
  for (const char of text) {
    width += CHAR_WIDTHS[char] ?? (/[0-9A-Z]/.test(char) ? 7 : 6.2);
  }
  return Math.ceil(width);
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export interface BadgeOptions {
  label: string;
  message: string;
  color: string;
  labelColor?: string;
  logo?: string;
}

/** A flat-square badge SVG (self-hosted, no external badge service). */
export function flatSquareBadge(options: BadgeOptions): string {
  const { label, message, color, labelColor = "#555" } = options;
  const logoPath = options.logo ? LOGOS[options.logo] : undefined;
  const pad = 10;
  const logoSpace = logoPath ? LOGO_SIZE + LOGO_GAP : 0;

  const labelWidth = textWidth(label) + pad + logoSpace;
  const messageWidth = textWidth(message) + pad;
  const width = labelWidth + messageWidth;
  const labelTextX = (logoSpace + labelWidth) / 2;
  const messageTextX = labelWidth + messageWidth / 2;

  const logo = logoPath
    ? `<g transform="translate(5 ${(20 - LOGO_SIZE) / 2}) scale(${LOGO_SIZE / 24})" fill="#fff"><path d="${logoPath}"/></g>`
    : "";

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="20" role="img" aria-label="${escapeXml(label)}: ${escapeXml(message)}">
<rect width="${labelWidth}" height="20" fill="${labelColor}"/>
<rect x="${labelWidth}" width="${messageWidth}" height="20" fill="${color}"/>
${logo}
<g fill="#fff" text-anchor="middle" font-family="Verdana,DejaVu Sans,Geneva,sans-serif" font-size="11">
<text x="${labelTextX}" y="14">${escapeXml(label)}</text>
<text x="${messageTextX}" y="14">${escapeXml(message)}</text>
</g>
</svg>`;
}
