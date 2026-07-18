/**
 * Text-edit command handlers — insert, replace, delete, format text,
 * format paragraph, apply style. Dispatched from executor.ts.
 */

import type { Document, Paragraph } from '../../types/document';
import type {
  InsertTextCommand,
  ReplaceTextCommand,
  DeleteTextCommand,
  FormatTextCommand,
  FormatParagraphCommand,
  ApplyStyleCommand,
} from '../../types/agentApi';
import {
  applyFormattingInParagraph,
  cloneDocument,
  deleteTextInParagraph,
  getBlockIndexForParagraph,
  getParagraphText,
  insertTextAtOffset,
} from './helpers';

/**
 * Insert text at a position
 */
export function executeInsertText(doc: Document, command: InsertTextCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;
  const blockIndex = getBlockIndexForParagraph(body, command.position.paragraphIndex);

  if (blockIndex === -1) {
    throw new Error(`Paragraph index ${command.position.paragraphIndex} not found`);
  }

  const paragraph = body.content[blockIndex] as Paragraph;
  paragraph.content = insertTextAtOffset(
    paragraph,
    command.position.offset,
    command.text,
    command.formatting
  );

  return newDoc;
}

/**
 * Replace text in a range
 */
export function executeReplaceText(doc: Document, command: ReplaceTextCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  const { start, end } = command.range;

  if (start.paragraphIndex === end.paragraphIndex) {
    // Same paragraph
    const blockIndex = getBlockIndexForParagraph(body, start.paragraphIndex);
    if (blockIndex === -1) {
      throw new Error(`Paragraph index ${start.paragraphIndex} not found`);
    }

    const paragraph = body.content[blockIndex] as Paragraph;

    // Delete the range first
    paragraph.content = deleteTextInParagraph(paragraph, start.offset, end.offset);

    // Then insert the new text
    paragraph.content = insertTextAtOffset(
      paragraph,
      start.offset,
      command.text,
      command.formatting
    );
  } else {
    // Multiple paragraphs - simplify by deleting and inserting
    // Delete from start to end of first paragraph
    const startBlockIndex = getBlockIndexForParagraph(body, start.paragraphIndex);
    const startParagraph = body.content[startBlockIndex] as Paragraph;
    const startText = getParagraphText(startParagraph);

    startParagraph.content = deleteTextInParagraph(startParagraph, start.offset, startText.length);
    startParagraph.content = insertTextAtOffset(
      startParagraph,
      start.offset,
      command.text,
      command.formatting
    );

    // Delete intermediate paragraphs and beginning of last paragraph
    const paragraphsToRemove: number[] = [];
    for (let i = start.paragraphIndex + 1; i <= end.paragraphIndex; i++) {
      paragraphsToRemove.push(getBlockIndexForParagraph(body, i));
    }

    // Remove in reverse order to preserve indices
    for (let i = paragraphsToRemove.length - 1; i >= 0; i--) {
      if (paragraphsToRemove[i] !== -1) {
        body.content.splice(paragraphsToRemove[i], 1);
      }
    }
  }

  return newDoc;
}

/**
 * Delete text in a range
 */
export function executeDeleteText(doc: Document, command: DeleteTextCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  const { start, end } = command.range;

  if (start.paragraphIndex === end.paragraphIndex) {
    // Same paragraph
    const blockIndex = getBlockIndexForParagraph(body, start.paragraphIndex);
    if (blockIndex === -1) {
      throw new Error(`Paragraph index ${start.paragraphIndex} not found`);
    }

    const paragraph = body.content[blockIndex] as Paragraph;
    paragraph.content = deleteTextInParagraph(paragraph, start.offset, end.offset);
  } else {
    // Multiple paragraphs
    // Truncate first paragraph
    const startBlockIndex = getBlockIndexForParagraph(body, start.paragraphIndex);
    const startParagraph = body.content[startBlockIndex] as Paragraph;
    const startText = getParagraphText(startParagraph);
    startParagraph.content = deleteTextInParagraph(startParagraph, start.offset, startText.length);

    // Delete intermediate paragraphs and truncate last
    const endBlockIndex = getBlockIndexForParagraph(body, end.paragraphIndex);
    const endParagraph = body.content[endBlockIndex] as Paragraph;
    endParagraph.content = deleteTextInParagraph(endParagraph, 0, end.offset);

    // Merge last paragraph content into first
    startParagraph.content.push(...endParagraph.content);

    // Remove paragraphs between start and end (inclusive of end)
    const indicesToRemove: number[] = [];
    for (let i = start.paragraphIndex + 1; i <= end.paragraphIndex; i++) {
      indicesToRemove.push(getBlockIndexForParagraph(body, i));
    }

    for (let i = indicesToRemove.length - 1; i >= 0; i--) {
      if (indicesToRemove[i] !== -1) {
        body.content.splice(indicesToRemove[i], 1);
      }
    }
  }

  return newDoc;
}

/**
 * Apply formatting to a range
 */
export function executeFormatText(doc: Document, command: FormatTextCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  const { start, end } = command.range;

  if (start.paragraphIndex === end.paragraphIndex) {
    // Same paragraph
    const blockIndex = getBlockIndexForParagraph(body, start.paragraphIndex);
    if (blockIndex === -1) {
      throw new Error(`Paragraph index ${start.paragraphIndex} not found`);
    }

    const paragraph = body.content[blockIndex] as Paragraph;
    paragraph.content = applyFormattingInParagraph(
      paragraph,
      start.offset,
      end.offset,
      command.formatting
    );
  } else {
    // Multiple paragraphs
    for (let i = start.paragraphIndex; i <= end.paragraphIndex; i++) {
      const blockIndex = getBlockIndexForParagraph(body, i);
      if (blockIndex === -1) continue;

      const paragraph = body.content[blockIndex] as Paragraph;
      const paragraphText = getParagraphText(paragraph);

      let startOffset = 0;
      let endOffset = paragraphText.length;

      if (i === start.paragraphIndex) {
        startOffset = start.offset;
      }
      if (i === end.paragraphIndex) {
        endOffset = end.offset;
      }

      paragraph.content = applyFormattingInParagraph(
        paragraph,
        startOffset,
        endOffset,
        command.formatting
      );
    }
  }

  return newDoc;
}

/**
 * Apply paragraph formatting
 */
export function executeFormatParagraph(doc: Document, command: FormatParagraphCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  const blockIndex = getBlockIndexForParagraph(body, command.paragraphIndex);
  if (blockIndex === -1) {
    throw new Error(`Paragraph index ${command.paragraphIndex} not found`);
  }

  const paragraph = body.content[blockIndex] as Paragraph;
  paragraph.formatting = { ...paragraph.formatting, ...command.formatting };

  // Handle listRendering when numPr changes
  if ('numPr' in command.formatting) {
    const numPr = command.formatting.numPr;
    if (numPr && numPr.numId !== undefined && numPr.numId !== 0) {
      // Setting a list - compute listRendering
      const ilvl = numPr.ilvl ?? 0;
      const isBullet = numPr.numId === 1; // numId 1 is typically bullets, 2 is numbered

      // Try to get marker from numbering definitions if available
      let marker = isBullet ? '•' : `${1}.`; // Default markers

      if (newDoc.package.numbering) {
        const num = newDoc.package.numbering.nums.find((n) => n.numId === numPr.numId);
        if (num) {
          const abstractNum = newDoc.package.numbering.abstractNums.find(
            (a) => a.abstractNumId === num.abstractNumId
          );
          if (abstractNum) {
            const level = abstractNum.levels.find((l) => l.ilvl === ilvl);
            if (level) {
              marker = level.lvlText || marker;
            }
          }
        }
      }

      paragraph.listRendering = {
        level: ilvl,
        numId: numPr.numId,
        marker,
        isBullet,
      };
    } else {
      // Removing list - clear listRendering
      delete paragraph.listRendering;
    }
  }

  return newDoc;
}

/**
 * Apply a named style to a paragraph
 */
export function executeApplyStyle(doc: Document, command: ApplyStyleCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  const blockIndex = getBlockIndexForParagraph(body, command.paragraphIndex);
  if (blockIndex === -1) {
    throw new Error(`Paragraph index ${command.paragraphIndex} not found`);
  }

  const paragraph = body.content[blockIndex] as Paragraph;
  paragraph.formatting = {
    ...paragraph.formatting,
    styleId: command.styleId,
  };

  return newDoc;
}
