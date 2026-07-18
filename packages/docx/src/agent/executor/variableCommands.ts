/**
 * Template-variable command handlers — register a variable and apply
 * `{variable}` substitutions across the document. Dispatched from
 * executor.ts.
 */

import type { Document, Paragraph, Run, BlockContent } from '../../types/document';
import type { SetVariableCommand, ApplyVariablesCommand } from '../../types/agentApi';
import { cloneDocument } from './helpers';

/**
 * Set a template variable value
 */
export function executeSetVariable(doc: Document, command: SetVariableCommand): Document {
  const newDoc = cloneDocument(doc);

  // Store variable in document for later application
  if (!newDoc.templateVariables) {
    newDoc.templateVariables = [];
  }

  if (!newDoc.templateVariables.includes(command.name)) {
    newDoc.templateVariables.push(command.name);
  }

  // Note: Actual variable substitution happens in applyVariables
  return newDoc;
}

/**
 * Apply all template variables
 */
export function executeApplyVariables(doc: Document, command: ApplyVariablesCommand): Document {
  const newDoc = cloneDocument(doc);
  const body = newDoc.package.document;

  // Replace {variable} patterns in all text content
  function replaceVariablesInRun(run: Run): void {
    for (const content of run.content) {
      if (content.type === 'text') {
        for (const [name, value] of Object.entries(command.values)) {
          const pattern = new RegExp(`\\{${name}\\}`, 'g');
          content.text = content.text.replace(pattern, value);
        }
      }
    }
  }

  function replaceVariablesInParagraph(paragraph: Paragraph): void {
    for (const item of paragraph.content) {
      if (item.type === 'run') {
        replaceVariablesInRun(item);
      } else if (item.type === 'hyperlink') {
        for (const child of item.children) {
          if (child.type === 'run') {
            replaceVariablesInRun(child);
          }
        }
      }
    }
  }

  function replaceVariablesInBlock(block: BlockContent): void {
    if (block.type === 'paragraph') {
      replaceVariablesInParagraph(block);
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          for (const cellBlock of cell.content) {
            replaceVariablesInBlock(cellBlock);
          }
        }
      }
    }
  }

  for (const block of body.content) {
    replaceVariablesInBlock(block);
  }

  return newDoc;
}
