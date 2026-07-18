/** Rust-backed selective DOCX save compatibility wrapper. */

import type { BlockContent, Document } from '../types/document';
import { writeDocumentWithRust } from './rustSaveFacade';

export interface SelectiveSaveOptions {
  changedParaIds: Set<string>;
  structuralChange: boolean;
  hasUntrackedChanges: boolean;
}

/**
 * Patch changed paragraphs through the Rust S13 writer. Returns `null` when
 * selective save is unsafe so callers can fall back to a full repack.
 */
export async function attemptSelectiveSave(
  doc: Document,
  originalBuffer: ArrayBuffer,
  options: SelectiveSaveOptions
): Promise<ArrayBuffer | null> {
  if (!canAttemptSelectiveSave(doc, originalBuffer, options)) return null;
  try {
    return (
      await writeDocumentWithRust(
        doc,
        originalBuffer,
        { updateModifiedDate: true },
        { changedParaIds: options.changedParaIds }
      )
    ).buffer;
  } catch {
    return null;
  }
}

function canAttemptSelectiveSave(
  doc: Document,
  originalBuffer: ArrayBuffer,
  options: SelectiveSaveOptions
): boolean {
  if (options.structuralChange || options.hasUntrackedChanges || !originalBuffer) return false;
  return !hasNewImagesOrHyperlinks(doc.package.document.content);
}

function hasNewImagesOrHyperlinks(blocks: BlockContent[]): boolean {
  const runHasNewImage = (run: {
    content: { type: string; image?: { src?: string; rId?: string } }[];
  }): boolean =>
    run.content.some(
      (content) =>
        content.type === 'drawing' &&
        content.image?.src?.startsWith('data:') &&
        !content.image.rId
    );

  for (const block of blocks) {
    if (block.type === 'paragraph') {
      for (const item of block.content) {
        if (item.type === 'run' && runHasNewImage(item)) return true;
        if (item.type === 'hyperlink' && item.href && !item.rId && !item.anchor) return true;
        if (
          item.type === 'insertion' ||
          item.type === 'deletion' ||
          item.type === 'moveFrom' ||
          item.type === 'moveTo'
        ) {
          for (const child of item.content) {
            if (child.type === 'run' && runHasNewImage(child)) return true;
          }
        }
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          if (hasNewImagesOrHyperlinks(cell.content)) return true;
        }
      }
    } else if (block.type === 'blockSdt' && hasNewImagesOrHyperlinks(block.content)) {
      return true;
    }
  }
  return false;
}
