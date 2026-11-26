import type { Document, HeaderFooter } from '../types/document';
import { isWrittenByCanonical } from './headerFooterAliasProjection';
import { preloadParseWasm } from './parseWasm';
import { parseDocumentWithRust } from './rustParseFacade';
import { assertSafeSaveTree } from './rustSaveFacade';
import { withSessionRevisionIds } from './sessionRevisionIds';
import { preloadOpcWasm, unzipContainer } from './wasm';

const MODEL_PARTS_REWRITTEN = 'DOCX_MODEL_PARTS_REWRITTEN';

function headerFooterSaveEntries(
  entries: Map<string, HeaderFooter> | undefined
): Array<[string, HeaderFooter]> {
  return entries
    ? [...entries].filter(([, part]) => !isWrittenByCanonical(part, entries))
    : [];
}

/** The model fields `writeDocumentWithRust` sends to the writer, serialized the same way. */
export function serializeDocumentSaveModel(document: Document): string {
  const pkg = document.package;
  const model = {
    document: pkg.document,
    headerEntries: headerFooterSaveEntries(pkg.headers),
    footerEntries: headerFooterSaveEntries(pkg.footers),
    footnotes: pkg.footnotes ?? [],
    endnotes: pkg.endnotes ?? [],
    footnoteSeparators: pkg.footnoteSeparators ?? [],
    endnoteSeparators: pkg.endnoteSeparators ?? [],
    relationshipEntries: [...(pkg.relationships?.entries() ?? [])],
    ...(pkg.numbering === undefined ? {} : { numbering: pkg.numbering }),
  };
  assertSafeSaveTree(model, 'save');
  return withSessionRevisionIds(model, () => JSON.stringify(model));
}

export async function modelMatchesSource(doc: Document, source: ArrayBuffer): Promise<boolean> {
  try {
    await Promise.all([preloadParseWasm(), preloadOpcWasm()]);
    const { document } = parseDocumentWithRust(source, {
      parseHeadersFooters: true,
      parseNotes: true,
      detectVariables: false,
    });
    return serializeDocumentSaveModel(doc) === serializeDocumentSaveModel(document);
  } catch {
    return false;
  }
}

export function modelPartsRewrittenWarning(source: ArrayBuffer, output: ArrayBuffer): string {
  let parts: string;
  try {
    const original = unzipContainer(new Uint8Array(source));
    const written = unzipContainer(new Uint8Array(output));
    const paths = Object.entries(written)
      .filter(([path, bytes]) => {
        const before = original[path];
        return path !== 'docProps/core.xml' && (!before ||
          before.byteLength !== bytes.byteLength || bytes.some((byte, i) => byte !== before[i]));
      })
      .map(([path]) => path)
      .sort();
    parts = JSON.stringify(paths);
  } catch {
    parts = '(part names unavailable)';
  }
  return `${MODEL_PARTS_REWRITTEN} ${parts} Use createYrsSession/seedFromDocx/saveYrsDocx to keep unchanged content when editing.`;
}
