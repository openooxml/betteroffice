/**
 * Headless structured export: DOCX bytes in, read-only structured content or Markdown out. The
 * Rust engine does the reading and anchors address the returned snapshot only. Plain exports
 * involve no session, editor, DOM or font; a paged export lays the bytes out in a private session
 * with exactly the fonts it is given, in a font store no other session shares.
 */

import type { Document } from '../types/document';
import type { YrsRenderEnv } from '../yrs';
import type {
  DocxHeadlessFont,
  DocxHeadlessLayoutOptions,
  DocxLayoutMap,
  DocxPageExportOptions,
  DocxPageMarkdownOptions,
  DocxPagedStructuredContent,
  DocxSnapshotLayoutMap,
} from '../yrs/pagedExport';
import type {
  DocxExportFailure,
  DocxExportOptions,
  DocxMarkdownContent,
  DocxMarkdownOptions,
  DocxStructuredContent,
} from '../yrs/structuredExport';

/** An export the engine refused as data: unusable options. */
export class DocxExportError extends Error {
  readonly failure: DocxExportFailure;

  constructor(failure: DocxExportFailure) {
    super(failure.message);
    this.name = 'DocxExportError';
    this.failure = failure;
  }
}

type Outcome<T> = { ok: true; content: T } | { ok: false; failure: DocxExportFailure };

async function edit() {
  const wasm = await import('../wasm/edit');
  await wasm.preloadEditWasm();
  return wasm;
}

function unwrap<T>(json: string): T {
  const outcome = JSON.parse(json) as Outcome<T>;
  if (!outcome.ok) throw new DocxExportError(outcome.failure);
  return outcome.content;
}

/**
 * Exports DOCX bytes as structured content. Throws {@link DocxExportError} for unusable options
 * and an `Error` for bytes that are not a readable DOCX.
 */
export async function exportDocxStructured(
  bytes: Uint8Array,
  options: DocxExportOptions
): Promise<DocxStructuredContent> {
  const wasm = await edit();
  return unwrap(wasm.exportDocxStructuredJson(bytes, JSON.stringify(options)));
}

/** {@link exportDocxStructured} rendered as Markdown with source anchors. */
export async function exportDocxMarkdown(
  bytes: Uint8Array,
  options: DocxExportOptions
): Promise<DocxMarkdownContent> {
  const wasm = await edit();
  return unwrap(wasm.exportDocxMarkdownJson(bytes, JSON.stringify(options)));
}

/** Renders structured content (schema version 1) as Markdown. */
export async function renderDocxMarkdown(
  content: DocxStructuredContent,
  options: DocxMarkdownOptions = {}
): Promise<DocxMarkdownContent> {
  const wasm = await edit();
  return unwrap(wasm.renderDocxMarkdownJson(JSON.stringify(content), JSON.stringify(options)));
}

/** The render environment the editor lays `document` out with, with `overrides` applied. */
function renderEnvironment(
  document: Document,
  overrides: DocxHeadlessLayoutOptions['renderEnvironment']
): YrsRenderEnv {
  const themeColors: Record<string, string> = {};
  for (const [name, value] of Object.entries(document.package.theme?.colorScheme ?? {})) {
    if (typeof value === 'string') themeColors[name] = value;
  }
  return {
    themeColors,
    defaultTabStopTwips:
      overrides?.defaultTabStopTwips ?? document.package.settings?.defaultTabStop ?? null,
    numericIds: {},
    showHiddenText: overrides?.showHiddenText ?? false,
  };
}

/** The bytes of a base64-encoded font. */
function fontBytes(font: DocxHeadlessFont): Uint8Array {
  let binary: string;
  try {
    binary = atob(font.data);
  } catch {
    throw new TypeError(`font ${JSON.stringify(font.key)} is not base64`);
  }
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

/**
 * Exports DOCX bytes as structured content with a page map. A private session opens the bytes
 * and lays every section, header, footer and note out with exactly `layout.fonts`, registered in
 * a measurement font store of its own so no editor's fonts change, and exports from that one
 * captured state. Anchors address the returned snapshot and the map is deterministic for the
 * same bytes, fonts and options. Throws {@link DocxExportError} for a refusal, a `TypeError` for
 * an unknown, repeated or undecodable font key and an `Error` for bytes that are not a readable
 * DOCX or fonts the engine rejects.
 */
export async function exportDocxStructuredWithPages(
  bytes: Uint8Array,
  options: DocxPageExportOptions,
  layout: DocxHeadlessLayoutOptions
): Promise<DocxPagedStructuredContent<DocxSnapshotLayoutMap>> {
  await edit();
  const [{ createYrsSession }, { buildResidentRegionLayoutRequest }] = await Promise.all([
    import('../yrs'),
    import('../editor/computeLayout'),
  ]);
  const indexes = new Map<string, number>();
  const files = layout.fonts.map((font, index) => {
    if (indexes.has(font.key)) throw new TypeError(`font key ${JSON.stringify(font.key)} is repeated`);
    indexes.set(font.key, index);
    return fontBytes(font);
  });
  const chain = (keys: readonly string[]): number[] =>
    keys.map((key) => {
      const index = indexes.get(key);
      if (index === undefined) throw new TypeError(`font key ${JSON.stringify(key)} is not a font`);
      return index;
    });
  const session = await createYrsSession({ clientId: 1 });
  try {
    const { document } = session.openDocx(bytes, true);
    const request = buildResidentRegionLayoutRequest(
      document,
      0,
      renderEnvironment(document, layout.renderEnvironment)
    );
    const compat = document.package.settings?.compatibilityFlags;
    const fontChains: Record<string, number[]> = {};
    request.measurement = {
      fontChains,
      defaults: {
        fontSize: layout.measurementDefaults?.fontSize ?? 11,
        fontFamily: layout.measurementDefaults?.fontFamily ?? 'Calibri',
      },
      compat: {
        noLeading: layout.compatibility?.noLeading ?? compat?.noLeading ?? false,
        doNotExpandShiftReturn:
          layout.compatibility?.doNotExpandShiftReturn ?? compat?.doNotExpandShiftReturn ?? false,
      },
      authoritativeShaping: true,
    };
    const requirements = JSON.parse(
      session.layoutFontRequirementsJson(JSON.stringify(request))
    ) as Array<{ key: string; family: string }>;
    for (const requirement of requirements) {
      fontChains[requirement.key] = chain(
        layout.fontChains?.[requirement.key] ??
          layout.fontChains?.[requirement.family.toLowerCase()] ??
          layout.defaultChain
      );
    }
    const joined = new Uint8Array(files.reduce((total, file) => total + file.length, 0));
    let offset = 0;
    for (const file of files) {
      joined.set(file, offset);
      offset += file.length;
    }
    const outcome = session.exportSnapshotWithPrivateFonts(
      joined,
      Uint32Array.from(files, (file) => file.length),
      JSON.stringify(request),
      options
    );
    if (!outcome.ok) throw new DocxExportError(outcome.failure);
    return outcome.content;
  } finally {
    session.destroy();
  }
}

/**
 * Renders a paged export as Markdown. With `pageMarkers`, each block marker is followed by
 * `<!-- docx-pages: N=label ... -->`: the physical page indexes and percent-encoded displayed
 * labels of every page showing the block. No page break is written into the text. Throws
 * {@link DocxExportError} when the page map was not captured with the content.
 */
export async function renderDocxMarkdownWithPages(
  content: DocxPagedStructuredContent<DocxLayoutMap | DocxSnapshotLayoutMap>,
  options: DocxPageMarkdownOptions = {}
): Promise<DocxMarkdownContent> {
  const wasm = await edit();
  return unwrap(
    wasm.renderDocxMarkdownWithPagesJson(JSON.stringify(content), JSON.stringify(options))
  );
}
