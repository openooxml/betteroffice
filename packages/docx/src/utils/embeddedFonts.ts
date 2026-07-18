/**
 * Embedded font loading.
 *
 * Turns the obfuscated `word/fonts/*.odttf` binaries a DOCX carries into live
 * `@font-face` registrations so documents render in their authored fonts
 * instead of falling back to a metric-compatible substitute.
 *
 * The pure half ({@link getEmbeddedFontFaces}) resolves + de-obfuscates the
 * faces; {@link loadEmbeddedFonts} registers them with the browser (no-op
 * outside a DOM). Round-trip preservation is unaffected: the serializer copies
 * `word/fonts/*` and `fontTable.xml` from the original package untouched.
 */

import { parseRelationshipsXmlWithRust } from '../docx/rustParseFacade';
import { readDocxContainer } from '../docx/zipContainer';
import { deobfuscateFont, isValidFontKey } from './fontDeobfuscation';
import { loadFontFromBuffer } from './fontLoader';
import type { FontTable, FontInfo, FontEmbed } from '../types/styles';
import type { Document } from '../types/document';

/**
 * Names of the fonts a table embeds at least one face for. Used by the picker
 * to surface embedded fonts even when the canvas probe is unreliable for a
 * subsetted face.
 *
 * @public
 */
export function getEmbeddedFontFamilies(fontTable: FontTable | undefined): Set<string> {
  const families = new Set<string>();
  if (!fontTable) return families;
  for (const font of fontTable.fonts) {
    if (font.embedRegular || font.embedBold || font.embedItalic || font.embedBoldItalic) {
      families.add(font.name);
    }
  }
  return families;
}

/** A single de-obfuscated embedded font face, ready for `loadFontFromBuffer`. */
export interface EmbeddedFontFace {
  /** Word font name to register the face under. */
  family: string;
  /** CSS `font-weight` the face maps to (`embedBold*` → `'bold'`). */
  weight: 'normal' | 'bold';
  /** CSS `font-style` the face maps to (`embed*Italic` → `'italic'`). */
  style: 'normal' | 'italic';
  /** De-obfuscated OpenType/TrueType bytes. */
  data: ArrayBuffer;
  /** Whether the source face was subsetted (`w:subsetted`). */
  subsetted: boolean;
}

interface EmbedKind {
  key: keyof Pick<FontInfo, 'embedRegular' | 'embedBold' | 'embedItalic' | 'embedBoldItalic'>;
  weight: 'normal' | 'bold';
  style: 'normal' | 'italic';
}

const EMBED_KINDS: EmbedKind[] = [
  { key: 'embedRegular', weight: 'normal', style: 'normal' },
  { key: 'embedBold', weight: 'bold', style: 'normal' },
  { key: 'embedItalic', weight: 'normal', style: 'italic' },
  { key: 'embedBoldItalic', weight: 'bold', style: 'italic' },
];

/**
 * Resolve a relationship target (e.g. `fonts/font1.odttf`, relative to the
 * `word/` folder) to the matching key in the unzipped font map, matched
 * case-insensitively (ZIP entries vary in case across producers).
 */
function lookupFontData(
  target: string,
  rawFonts: ReadonlyMap<string, ArrayBuffer>
): ArrayBuffer | undefined {
  const cleaned = target.replace(/^\/+/, '');
  const fullPath = cleaned.toLowerCase().startsWith('word/') ? cleaned : `word/${cleaned}`;
  const wanted = fullPath.toLowerCase();
  for (const [path, data] of rawFonts) {
    if (path.toLowerCase() === wanted) return data;
  }
  return undefined;
}

function resolveFace(
  family: string,
  embed: FontEmbed,
  kind: EmbedKind,
  rawFonts: ReadonlyMap<string, ArrayBuffer>,
  relTargets: ReadonlyMap<string, string>
): EmbeddedFontFace | null {
  const target = relTargets.get(embed.relId);
  if (!target) return null;

  const raw = lookupFontData(target, rawFonts);
  if (!raw) return null;

  let data: ArrayBuffer;
  if (embed.fontKey && isValidFontKey(embed.fontKey)) {
    try {
      data = deobfuscateFont(raw, embed.fontKey);
    } catch {
      return null;
    }
  } else {
    // No (valid) key: assume the bytes are already a usable font.
    data = raw.slice(0);
  }

  return {
    family,
    weight: kind.weight,
    style: kind.style,
    data,
    subsetted: embed.subsetted ?? false,
  };
}

/**
 * Resolve and de-obfuscate every embedded font face declared in a font table.
 * Pure: does not touch the DOM. Faces whose relationship or binary is missing,
 * or whose key is unusable, are skipped.
 *
 * @param fontTable - Parsed `fontTable.xml` (`pkg.fontTable`).
 * @param rawFonts - Unzipped font binaries keyed by package path.
 * @param fontTableRelsXml - Raw `word/_rels/fontTable.xml.rels` XML.
 * @public
 */
export function getEmbeddedFontFaces(
  fontTable: FontTable | undefined,
  rawFonts: ReadonlyMap<string, ArrayBuffer>,
  fontTableRelsXml: string | null | undefined
): EmbeddedFontFace[] {
  if (!fontTable || fontTable.fonts.length === 0) return [];
  if (!fontTableRelsXml || rawFonts.size === 0) return [];

  const rels = parseRelationshipsXmlWithRust(
    fontTableRelsXml,
    'word/_rels/fontTable.xml.rels'
  );
  const relTargets = new Map<string, string>();
  for (const [id, rel] of rels) relTargets.set(id, rel.target);

  const faces: EmbeddedFontFace[] = [];
  for (const font of fontTable.fonts) {
    for (const kind of EMBED_KINDS) {
      const embed = font[kind.key];
      if (!embed) continue;
      const face = resolveFace(font.name, embed, kind, rawFonts, relTargets);
      if (face) faces.push(face);
    }
  }
  return faces;
}

/**
 * Resolve + de-obfuscate the embedded faces of an already-parsed Document by
 * re-reading its `originalBuffer` (parseDocx registers the faces with the
 * browser during load but does not retain the raw bytes on the model).
 * Returns `[]` when the document declares no embeds, carries no original
 * buffer, or the parts cannot be resolved. The zip pass goes through the
 * hardened `unzipDocx` container facade.
 *
 * Used by the experimental Rust measurement source, which needs the byte
 * form of every embedded face for the wasm FontStore.
 *
 * @public
 */
export async function extractEmbeddedFontFaces(document: Document): Promise<EmbeddedFontFace[]> {
  const fontTable = document.package.fontTable;
  if (!fontTable || getEmbeddedFontFamilies(fontTable).size === 0) return [];
  const buffer = document.originalBuffer;
  if (!buffer) return [];

  const container = readDocxContainer(buffer);
  const rawFonts = new Map<string, ArrayBuffer>();
  let fontTableRelsXml: string | null = null;
  for (const path of container.paths()) {
    if (path.toLowerCase().startsWith('word/fonts/')) {
      const bytes = container.file(path);
      if (bytes) rawFonts.set(path, bytes.slice().buffer);
    }
    if (path.toLowerCase() === 'word/_rels/fonttable.xml.rels') {
      fontTableRelsXml = container.text(path);
    }
  }
  return getEmbeddedFontFaces(fontTable, rawFonts, fontTableRelsXml);
}

/**
 * Register every embedded font face with the browser via `@font-face`. No-op
 * outside a DOM (headless/SSR). Resolves to the set of font family names that
 * were registered (deduped), so callers can surface them in the font picker.
 *
 * @public
 */
export async function loadEmbeddedFonts(
  fontTable: FontTable | undefined,
  rawFonts: ReadonlyMap<string, ArrayBuffer>,
  fontTableRelsXml: string | null | undefined
): Promise<Set<string>> {
  const faces = getEmbeddedFontFaces(fontTable, rawFonts, fontTableRelsXml);
  const families = new Set<string>();
  if (faces.length === 0) return families;

  await Promise.all(
    faces.map(async (face) => {
      const ok = await loadFontFromBuffer(face.family, face.data, {
        weight: face.weight,
        style: face.style,
      });
      if (ok) {
        families.add(face.family);
      } else if (typeof document !== 'undefined') {
        // In a browser, a false result means the face genuinely failed to load
        // (corrupt bytes, timeout) rather than headless no-op. Surface it: the
        // picker may still list this family (see getEmbeddedFontFamilies), and
        // text using it will fall back to the CSS stack.
        console.warn(
          `[embeddedFonts] embedded face "${face.family}" (${face.weight} ${face.style}) ` +
            'failed to load; text using it falls back to the CSS stack'
        );
      }
    })
  );
  return families;
}
