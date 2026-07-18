/**
 * Command Executor
 *
 * Executes agent commands on a document immutably:
 * - Handles all command types from AgentCommand
 * - Preserves surrounding formatting
 * - Returns new document (immutable updates)
 *
 * Per-command handlers live under `executor/` by domain (text edits,
 * paragraph splits/merges, structural inserts, template variables).
 * Shared helpers (clone, lookup, text-range edits) live in
 * `executor/helpers.ts`.
 */

import type { Document } from '../types/document';
import type { AgentCommand } from '../types/agentApi';

import { pluginRegistry } from '../core-plugins/registry';
import {
  executeInsertText,
  executeReplaceText,
  executeDeleteText,
  executeFormatText,
  executeFormatParagraph,
  executeApplyStyle,
} from './executor/textCommands';
import {
  executeInsertTable,
  executeInsertImage,
  executeInsertHyperlink,
  executeRemoveHyperlink,
} from './executor/structureCommands';
import {
  executeInsertParagraphBreak,
  executeMergeParagraphs,
  executeSplitParagraph,
} from './executor/paragraphCommands';
import { executeSetVariable, executeApplyVariables } from './executor/variableCommands';

/**
 * Execute an agent command on a document
 * Returns a new document with the command applied (immutable)
 *
 * Dispatch order:
 * 1. Try plugin handlers first (allows plugins to override built-in commands)
 * 2. Fall back to built-in handlers
 *
 * @param doc - The document to modify
 * @param command - The command to execute
 * @returns New document with command applied
 */
export function executeCommand(doc: Document, command: AgentCommand): Document {
  // Try plugin handlers first
  const pluginHandler = pluginRegistry.getCommandHandler(command.type);
  if (pluginHandler) {
    // Plugin commands use a more flexible type
    return pluginHandler(doc, command as unknown as import('../core-plugins/types').PluginCommand);
  }

  // Fall back to built-in handlers
  switch (command.type) {
    case 'insertText':
      return executeInsertText(doc, command);
    case 'replaceText':
      return executeReplaceText(doc, command);
    case 'deleteText':
      return executeDeleteText(doc, command);
    case 'formatText':
      return executeFormatText(doc, command);
    case 'formatParagraph':
      return executeFormatParagraph(doc, command);
    case 'applyStyle':
      return executeApplyStyle(doc, command);
    case 'insertTable':
      return executeInsertTable(doc, command);
    case 'insertImage':
      return executeInsertImage(doc, command);
    case 'insertHyperlink':
      return executeInsertHyperlink(doc, command);
    case 'removeHyperlink':
      return executeRemoveHyperlink(doc, command);
    case 'insertParagraphBreak':
      return executeInsertParagraphBreak(doc, command);
    case 'mergeParagraphs':
      return executeMergeParagraphs(doc, command);
    case 'splitParagraph':
      return executeSplitParagraph(doc, command);
    case 'setVariable':
      return executeSetVariable(doc, command);
    case 'applyVariables':
      return executeApplyVariables(doc, command);
    default: {
      // Exhaustive check — should never happen with proper types.
      const _exhaustive: never = command;
      throw new Error(`Unknown command type: ${(_exhaustive as AgentCommand).type}`);
    }
  }
}

/**
 * Execute multiple commands in sequence
 *
 * @param doc - The document to modify
 * @param commands - Commands to execute in order
 * @returns New document with all commands applied
 */
export function executeCommands(doc: Document, commands: AgentCommand[]): Document {
  return commands.reduce((currentDoc, command) => executeCommand(currentDoc, command), doc);
}
