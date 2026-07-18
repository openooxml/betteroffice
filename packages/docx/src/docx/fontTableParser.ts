/**
 * Font Table Parser
 *
 * Parses `word/fontTable.xml` into a {@link FontTable}. The main payload we
 * care about beyond font names is the `w:embed*` elements, which reference the
 * obfuscated font binaries under `word/fonts/` (resolved via the font table's
 * own `.rels`) and carry the `w:fontKey` GUID used to de-obfuscate them.
 *
 * Structure:
 * ```xml
 * <w:fonts>
 *   <w:font w:name="My Brand Sans">
 *     <w:family w:val="swiss"/>
 *     <w:pitch w:val="variable"/>
 *     <w:embedRegular r:id="rId1" w:fontKey="{GUID}" w:subsetted="true"/>
 *   </w:font>
 * </w:fonts>
 * ```
 */

import { parseXmlDocument, findChildren, findChild, getAttribute } from './xmlParser';
import type { FontTable, FontInfo, FontEmbed } from '../types/styles';
import type { XmlElement } from './xmlParser';

const FAMILY_VALUES = new Set(['decorative', 'modern', 'roman', 'script', 'swiss', 'auto']);
const PITCH_VALUES = new Set(['default', 'fixed', 'variable']);

function parseEmbed(font: XmlElement, localName: string): FontEmbed | undefined {
  const el = findChild(font, 'w', localName);
  if (!el) return undefined;

  const relId = getAttribute(el, 'r', 'id');
  if (!relId) return undefined; // an embed with no relationship is unusable

  const embed: FontEmbed = { relId };

  const fontKey = getAttribute(el, 'w', 'fontKey');
  if (fontKey) embed.fontKey = fontKey;

  const subsetted = getAttribute(el, 'w', 'subsetted');
  if (subsetted === 'true' || subsetted === '1') embed.subsetted = true;

  return embed;
}

function parseFontInfo(font: XmlElement): FontInfo | null {
  const name = getAttribute(font, 'w', 'name');
  if (!name) return null;

  const info: FontInfo = { name };

  const altName = findChild(font, 'w', 'altName');
  if (altName) {
    const val = getAttribute(altName, 'w', 'val');
    if (val) info.altName = val;
  }

  const family = findChild(font, 'w', 'family');
  if (family) {
    const val = getAttribute(family, 'w', 'val');
    if (val && FAMILY_VALUES.has(val)) info.family = val as FontInfo['family'];
  }

  const pitch = findChild(font, 'w', 'pitch');
  if (pitch) {
    const val = getAttribute(pitch, 'w', 'val');
    if (val && PITCH_VALUES.has(val)) info.pitch = val as FontInfo['pitch'];
  }

  const charset = findChild(font, 'w', 'charset');
  if (charset) {
    const val = getAttribute(charset, 'w', 'val');
    if (val) info.charset = val;
  }

  const panose1 = findChild(font, 'w', 'panose1');
  if (panose1) {
    const val = getAttribute(panose1, 'w', 'val');
    if (val) info.panose1 = val;
  }

  const embedRegular = parseEmbed(font, 'embedRegular');
  if (embedRegular) info.embedRegular = embedRegular;
  const embedBold = parseEmbed(font, 'embedBold');
  if (embedBold) info.embedBold = embedBold;
  const embedItalic = parseEmbed(font, 'embedItalic');
  if (embedItalic) info.embedItalic = embedItalic;
  const embedBoldItalic = parseEmbed(font, 'embedBoldItalic');
  if (embedBoldItalic) info.embedBoldItalic = embedBoldItalic;

  return info;
}

/**
 * Parse `word/fontTable.xml` into a {@link FontTable}. Returns an empty table
 * for missing/empty/unparseable input.
 *
 * @public
 */
export function parseFontTable(fontTableXml: string | null | undefined): FontTable {
  if (!fontTableXml || fontTableXml.trim().length === 0) {
    return { fonts: [] };
  }

  const root = parseXmlDocument(fontTableXml);
  if (!root) {
    console.warn('Failed to parse fontTable.xml');
    return { fonts: [] };
  }

  const fonts: FontInfo[] = [];
  for (const font of findChildren(root, 'w', 'font')) {
    const info = parseFontInfo(font);
    if (info) fonts.push(info);
  }

  return { fonts };
}
