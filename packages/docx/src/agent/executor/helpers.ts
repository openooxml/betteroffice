/**
 * Shared command-executor helpers.
 *
 * Pure functions for traversing and rewriting Document content:
 * paragraph lookup, plain-text extraction, immutable text edits.
 */

import type {
  Document,
  DocumentBody,
  Paragraph,
  Run,
  TextContent,
  ParagraphContent,
  TextFormatting,
} from '../../types/document';

/**
 * Deep clone a document for immutable updates
 */
export function cloneDocument(doc: Document): Document {
  return JSON.parse(JSON.stringify(doc));
}

/**
 * Get the block index for a paragraph index
 */
export function getBlockIndexForParagraph(body: DocumentBody, paragraphIndex: number): number {
  let currentParagraphIndex = 0;
  for (let i = 0; i < body.content.length; i++) {
    if (body.content[i].type === 'paragraph') {
      if (currentParagraphIndex === paragraphIndex) {
        return i;
      }
      currentParagraphIndex++;
    }
  }
  return -1;
}

/**
 * Get plain text from a paragraph
 */
export function getParagraphText(paragraph: Paragraph): string {
  let text = '';
  for (const item of paragraph.content) {
    if (item.type === 'run') {
      for (const content of item.content) {
        if (content.type === 'text') {
          text += content.text;
        }
      }
    } else if (item.type === 'hyperlink') {
      for (const child of item.children) {
        if (child.type === 'run') {
          for (const content of child.content) {
            if (content.type === 'text') {
              text += content.text;
            }
          }
        }
      }
    }
  }
  return text;
}

/**
 * Create a new run with text
 */
export function createTextRun(text: string, formatting?: TextFormatting): Run {
  return {
    type: 'run',
    formatting,
    content: [
      {
        type: 'text',
        text,
      },
    ],
  };
}

/**
 * Insert text at a specific offset within a paragraph
 * Returns new paragraph content
 */
export function insertTextAtOffset(
  paragraph: Paragraph,
  offset: number,
  text: string,
  formatting?: TextFormatting
): ParagraphContent[] {
  const newContent: ParagraphContent[] = [];
  let currentOffset = 0;
  let inserted = false;

  for (const item of paragraph.content) {
    if (item.type === 'run') {
      const runText = item.content
        .filter((c): c is TextContent => c.type === 'text')
        .map((c) => c.text)
        .join('');

      const runStart = currentOffset;
      const runEnd = currentOffset + runText.length;

      if (!inserted && offset >= runStart && offset <= runEnd) {
        // Insert within this run
        const insertPos = offset - runStart;

        if (insertPos > 0) {
          // Text before insertion point
          newContent.push({
            ...item,
            content: [{ type: 'text', text: runText.slice(0, insertPos) }],
          });
        }

        // New text
        newContent.push(createTextRun(text, formatting || item.formatting));

        if (insertPos < runText.length) {
          // Text after insertion point
          newContent.push({
            ...item,
            content: [{ type: 'text', text: runText.slice(insertPos) }],
          });
        }

        inserted = true;
      } else {
        newContent.push(item);
      }

      currentOffset = runEnd;
    } else {
      newContent.push(item);
    }
  }

  // If not inserted yet, append at the end
  if (!inserted) {
    newContent.push(createTextRun(text, formatting));
  }

  return newContent;
}

/**
 * Delete text in a range within a single paragraph
 */
export function deleteTextInParagraph(
  paragraph: Paragraph,
  startOffset: number,
  endOffset: number
): ParagraphContent[] {
  const newContent: ParagraphContent[] = [];
  let currentOffset = 0;

  for (const item of paragraph.content) {
    if (item.type === 'run') {
      const runText = item.content
        .filter((c): c is TextContent => c.type === 'text')
        .map((c) => c.text)
        .join('');

      const runStart = currentOffset;
      const runEnd = currentOffset + runText.length;

      // Check if run overlaps with deletion range
      if (runEnd <= startOffset || runStart >= endOffset) {
        // No overlap, keep entire run
        newContent.push(item);
      } else {
        // Partial overlap
        let newText = '';

        if (runStart < startOffset) {
          // Keep text before start
          newText += runText.slice(0, startOffset - runStart);
        }

        if (runEnd > endOffset) {
          // Keep text after end
          newText += runText.slice(endOffset - runStart);
        }

        if (newText.length > 0) {
          newContent.push({
            ...item,
            content: [{ type: 'text', text: newText }],
          });
        }
      }

      currentOffset = runEnd;
    } else {
      newContent.push(item);
    }
  }

  return newContent;
}

/**
 * Apply formatting to text in a range within a paragraph
 */
export function applyFormattingInParagraph(
  paragraph: Paragraph,
  startOffset: number,
  endOffset: number,
  formatting: Partial<TextFormatting>
): ParagraphContent[] {
  const newContent: ParagraphContent[] = [];
  let currentOffset = 0;

  for (const item of paragraph.content) {
    if (item.type === 'run') {
      const runText = item.content
        .filter((c): c is TextContent => c.type === 'text')
        .map((c) => c.text)
        .join('');

      const runStart = currentOffset;
      const runEnd = currentOffset + runText.length;

      // Check if run overlaps with formatting range
      if (runEnd <= startOffset || runStart >= endOffset) {
        // No overlap, keep entire run unchanged
        newContent.push(item);
      } else if (runStart >= startOffset && runEnd <= endOffset) {
        // Entire run is within range, apply formatting
        newContent.push({
          ...item,
          formatting: { ...item.formatting, ...formatting },
        });
      } else {
        // Partial overlap - need to split run
        const overlapStart = Math.max(startOffset, runStart);
        const overlapEnd = Math.min(endOffset, runEnd);

        // Text before overlap
        if (runStart < overlapStart) {
          newContent.push({
            ...item,
            content: [{ type: 'text', text: runText.slice(0, overlapStart - runStart) }],
          });
        }

        // Overlapping text with formatting
        newContent.push({
          ...item,
          formatting: { ...item.formatting, ...formatting },
          content: [
            {
              type: 'text',
              text: runText.slice(overlapStart - runStart, overlapEnd - runStart),
            },
          ],
        });

        // Text after overlap
        if (runEnd > overlapEnd) {
          newContent.push({
            ...item,
            content: [{ type: 'text', text: runText.slice(overlapEnd - runStart) }],
          });
        }
      }

      currentOffset = runEnd;
    } else {
      newContent.push(item);
    }
  }

  return newContent;
}
