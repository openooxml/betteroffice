/**
 * Paragraph-structure command handlers — insert paragraph break, merge
 * paragraphs, split paragraph. Dispatched from executor.ts.
 */

import type { Document, Paragraph } from '../../types/document';
import type {
  InsertParagraphBreakCommand,
  MergeParagraphsCommand,
  SplitParagraphCommand,
} from '../../types/agentApi';
import {
  cloneDocument,
  deleteTextInParagraph,
  getBlockIndexForParagraph,
  getParagraphText,
} from './helpers';

/**
 * Insert a paragraph break
 */
export function executeInsertParagraphBreak(
  doc: Document,
  command: InsertParagraphBreakCommand
): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  const blockIndex = getBlockIndexForParagraph(body, command.position.paragraphIndex);
  if (blockIndex === -1) {
    throw new Error(`Paragraph index ${command.position.paragraphIndex} not found`);
  }

  const paragraph = body.content[blockIndex] as Paragraph;
  const paragraphText = getParagraphText(paragraph);

  // Split the paragraph at the offset
  const beforeContent = deleteTextInParagraph(
    { ...paragraph, content: [...paragraph.content] },
    command.position.offset,
    paragraphText.length
  );

  const afterContent = deleteTextInParagraph(
    { ...paragraph, content: [...paragraph.content] },
    0,
    command.position.offset
  );

  // Update current paragraph with content before break
  paragraph.content = beforeContent;

  // Create new paragraph with content after break
  const newParagraph: Paragraph = {
    type: 'paragraph',
    formatting: paragraph.formatting,
    content: afterContent,
  };

  // Insert new paragraph after current one
  body.content.splice(blockIndex + 1, 0, newParagraph);

  return newDoc;
}

/**
 * Merge paragraphs
 */
export function executeMergeParagraphs(doc: Document, command: MergeParagraphsCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  const startBlockIndex = getBlockIndexForParagraph(body, command.paragraphIndex);
  if (startBlockIndex === -1) {
    throw new Error(`Paragraph index ${command.paragraphIndex} not found`);
  }

  const baseParagraph = body.content[startBlockIndex] as Paragraph;

  // Collect all content from paragraphs to merge
  const indicesToRemove: number[] = [];

  for (let i = 1; i <= command.count; i++) {
    const blockIndex = getBlockIndexForParagraph(body, command.paragraphIndex + i);
    if (blockIndex !== -1) {
      const para = body.content[blockIndex] as Paragraph;
      baseParagraph.content.push(...para.content);
      indicesToRemove.push(blockIndex);
    }
  }

  // Remove merged paragraphs in reverse order
  for (let i = indicesToRemove.length - 1; i >= 0; i--) {
    body.content.splice(indicesToRemove[i], 1);
  }

  return newDoc;
}

/**
 * Split a paragraph at a position
 */
export function executeSplitParagraph(doc: Document, command: SplitParagraphCommand): Document {
  // Split is the same as insert paragraph break
  return executeInsertParagraphBreak(doc, {
    type: 'insertParagraphBreak',
    position: command.position,
  });
}
