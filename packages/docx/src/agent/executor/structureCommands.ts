/**
 * Structural-insert command handlers — insert table, insert image, insert
 * and remove hyperlink. Dispatched from executor.ts.
 */

import type {
  Document,
  Paragraph,
  Run,
  Table,
  TableCell,
  TableRow,
  TextContent,
  ParagraphContent,
  Image,
  Hyperlink,
} from '../../types/document';
import type {
  InsertTableCommand,
  InsertImageCommand,
  InsertHyperlinkCommand,
  RemoveHyperlinkCommand,
} from '../../types/agentApi';
import { pixelsToEmu } from '../../utils/units';
import { DEFAULT_TABLE_WIDTH_DXA, distributeColumnWidths } from '../../utils/tableWidth';
import {
  createStyleResolver,
  pickTableStyleForInsert,
  DEFAULT_NEW_TABLE_LOOK,
} from '../../styles';
import {
  cloneDocument,
  createTextRun,
  deleteTextInParagraph,
  getBlockIndexForParagraph,
  getParagraphText,
  insertTextAtOffset,
} from './helpers';

/**
 * Insert a table at a position
 */
export function executeInsertTable(doc: Document, command: InsertTableCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;
  const columnWidths = distributeColumnWidths(DEFAULT_TABLE_WIDTH_DXA, command.columns);

  // Create table structure
  const rows: TableRow[] = [];

  for (let r = 0; r < command.rows; r++) {
    const cells: TableCell[] = [];

    for (let c = 0; c < command.columns; c++) {
      const cellText = command.data?.[r]?.[c] || '';
      cells.push({
        type: 'tableCell',
        formatting: columnWidths[c]
          ? { width: { value: columnWidths[c], type: 'dxa' } }
          : undefined,
        content: [
          {
            type: 'paragraph',
            content: cellText ? [createTextRun(cellText)] : [],
          },
        ],
      });
    }

    rows.push({
      type: 'tableRow',
      formatting: r === 0 && command.hasHeader ? { header: true } : undefined,
      cells,
    });
  }

  // Adopt the document's default table style when it declares one (settings
  // `w:defaultTableStyle`, else the type-default table style) so agent-inserted
  // tables match the document/template, just like the toolbar insert. The
  // styleId serializes as `<w:tblStyle>` and re-resolves through the same
  // OOXML cascade (`convertTable`) on render. When no usable default style
  // exists, the table carries no styleId and inherits the type-default style's
  // borders via that cascade, preserving prior behavior.
  const resolver = createStyleResolver(newDoc.package.styles);
  const preferredStyleId = pickTableStyleForInsert(
    newDoc.package.settings?.defaultTableStyle,
    resolver
  );

  // Give agent-inserted tables an explicit fixed grid so they serialize with a
  // valid w:tblGrid and render with predictable column widths.
  const table: Table = {
    type: 'table',
    formatting: {
      width: { value: DEFAULT_TABLE_WIDTH_DXA, type: 'dxa' },
      layout: 'fixed',
      ...(preferredStyleId
        ? { styleId: preferredStyleId, look: { ...DEFAULT_NEW_TABLE_LOOK } }
        : {}),
    },
    columnWidths,
    rows,
  };

  // Insert table after the specified paragraph
  const blockIndex = getBlockIndexForParagraph(body, command.position.paragraphIndex);
  if (blockIndex === -1) {
    body.content.push(table);
  } else {
    body.content.splice(blockIndex + 1, 0, table);
  }

  return newDoc;
}

/**
 * Insert an image at a position
 */
export function executeInsertImage(doc: Document, command: InsertImageCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  const blockIndex = getBlockIndexForParagraph(body, command.position.paragraphIndex);
  if (blockIndex === -1) {
    throw new Error(`Paragraph index ${command.position.paragraphIndex} not found`);
  }

  const paragraph = body.content[blockIndex] as Paragraph;

  // Create image
  const image: Image = {
    type: 'image',
    rId: `rId_img_${Date.now()}`,
    src: command.src,
    alt: command.alt,
    size: {
      width: pixelsToEmu(command.width || 100),
      height: pixelsToEmu(command.height || 100),
    },
    wrap: { type: 'inline' },
  };

  // Create run with drawing content
  const imageRun: Run = {
    type: 'run',
    content: [
      {
        type: 'drawing',
        image,
      },
    ],
  };

  // Insert image run at offset
  const newContent = insertTextAtOffset(paragraph, command.position.offset, '', undefined);
  // Find insertion point and add image
  let inserted = false;
  let currentOffset = 0;

  for (let i = 0; i < newContent.length; i++) {
    const item = newContent[i];
    if (item.type === 'run') {
      const runText = item.content
        .filter((c): c is TextContent => c.type === 'text')
        .map((c) => c.text)
        .join('');
      currentOffset += runText.length;

      if (!inserted && currentOffset >= command.position.offset) {
        newContent.splice(i + 1, 0, imageRun);
        inserted = true;
        break;
      }
    }
  }

  if (!inserted) {
    newContent.push(imageRun);
  }

  paragraph.content = newContent;

  return newDoc;
}

/**
 * Insert a hyperlink at a range
 */
export function executeInsertHyperlink(doc: Document, command: InsertHyperlinkCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  const { start, end } = command.range;

  if (start.paragraphIndex !== end.paragraphIndex) {
    throw new Error('Hyperlinks cannot span multiple paragraphs');
  }

  const blockIndex = getBlockIndexForParagraph(body, start.paragraphIndex);
  if (blockIndex === -1) {
    throw new Error(`Paragraph index ${start.paragraphIndex} not found`);
  }

  const paragraph = body.content[blockIndex] as Paragraph;
  const paragraphText = getParagraphText(paragraph);

  // Get the text that will become the link
  const linkText = command.displayText || paragraphText.slice(start.offset, end.offset);

  // Delete the original text
  paragraph.content = deleteTextInParagraph(paragraph, start.offset, end.offset);

  // Create hyperlink
  const hyperlink: Hyperlink = {
    type: 'hyperlink',
    href: command.url,
    tooltip: command.tooltip,
    children: [createTextRun(linkText)],
  };

  // Insert hyperlink at position
  let inserted = false;
  let currentOffset = 0;
  const newContent: ParagraphContent[] = [];

  for (const item of paragraph.content) {
    if (item.type === 'run') {
      const runText = item.content
        .filter((c): c is TextContent => c.type === 'text')
        .map((c) => c.text)
        .join('');

      const runEnd = currentOffset + runText.length;

      if (!inserted && currentOffset <= start.offset && start.offset <= runEnd) {
        const insertPos = start.offset - currentOffset;

        if (insertPos > 0) {
          newContent.push({
            ...item,
            content: [{ type: 'text', text: runText.slice(0, insertPos) }],
          });
        }

        newContent.push(hyperlink);

        if (insertPos < runText.length) {
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

  if (!inserted) {
    newContent.push(hyperlink);
  }

  paragraph.content = newContent;

  return newDoc;
}

/**
 * Remove a hyperlink but keep the text
 */
export function executeRemoveHyperlink(doc: Document, command: RemoveHyperlinkCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  const { start } = command.range;

  const blockIndex = getBlockIndexForParagraph(body, start.paragraphIndex);
  if (blockIndex === -1) {
    throw new Error(`Paragraph index ${start.paragraphIndex} not found`);
  }

  const paragraph = body.content[blockIndex] as Paragraph;
  const newContent: ParagraphContent[] = [];

  for (const item of paragraph.content) {
    if (item.type === 'hyperlink') {
      // Convert hyperlink children to regular runs
      for (const child of item.children) {
        if (child.type === 'run') {
          newContent.push(child);
        }
      }
    } else {
      newContent.push(item);
    }
  }

  paragraph.content = newContent;

  return newDoc;
}
