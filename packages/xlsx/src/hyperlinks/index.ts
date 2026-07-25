import type { DisplayList, HyperlinkRegion } from '../display-list/types';

export interface HyperlinkDestination {
  sheetName: string;
  row: number;
  col: number;
}

export function hyperlinkAtCell(
  displayList: DisplayList,
  row: number,
  col: number
): HyperlinkRegion | null {
  return (
    displayList.hyperlinks?.find(
      (link) =>
        row >= link.top && row <= link.bottom && col >= link.left && col <= link.right
    ) ?? null
  );
}

export function safeExternalHyperlink(link: HyperlinkRegion): string | null {
  if (!link.externalTarget) return null;
  let url: URL;
  try {
    url = new URL(link.externalTarget);
  } catch {
    return null;
  }
  if (!['http:', 'https:', 'mailto:', 'tel:'].includes(url.protocol)) return null;
  if (link.location && !url.hash) url.hash = link.location.replace(/^#/, '');
  return url.href;
}

export function parseHyperlinkLocation(
  location: string,
  currentSheet: string
): HyperlinkDestination | null {
  const target = location.replace(/^#/, '');
  const separator = target.lastIndexOf('!');
  const rawSheet = separator < 0 ? currentSheet : target.slice(0, separator);
  const rawAddress = (separator < 0 ? target : target.slice(separator + 1)).split(':', 1)[0];
  const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d*)$/.exec(rawAddress);
  if (!match) return null;
  let col = 0;
  for (const character of match[1].toUpperCase()) {
    col = col * 26 + character.charCodeAt(0) - 64;
  }
  const row = Number(match[2]);
  if (col < 1 || col > 16_384 || row < 1 || row > 1_048_576) return null;
  const sheetName =
    rawSheet.startsWith("'") && rawSheet.endsWith("'")
      ? rawSheet.slice(1, -1).replace(/''/g, "'")
      : rawSheet;
  if (!sheetName) return null;
  return { sheetName, row: row - 1, col: col - 1 };
}
