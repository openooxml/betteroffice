/**
 * Document Agent
 *
 * Headless, framework-agnostic API for inspecting and editing the
 * Document model for adapters that want agent capabilities without UI.
 * @packageDocumentation
 * @public
 */

export { DocumentAgent, createAgent, createAgentFromDocument } from './DocumentAgent';
export type {
  InsertTextOptions,
  InsertTableOptions,
  InsertImageOptions,
  InsertHyperlinkOptions,
  FormattedTextSegment,
} from './DocumentAgent';

export { executeCommand, executeCommands } from './executor';

export {
  getAgentContext,
  getDocumentSummary,
  buildSelectionContext as buildSelectionContextFromContext,
  type AgentContextOptions,
  type SelectionContextOptions as ContextSelectionOptions,
} from './context';

export {
  buildSelectionContext,
  buildExtendedSelectionContext,
  getSelectionFormattingSummary,
  type SelectionContextOptions,
  type ExtendedSelectionContext,
  type FormattingSummary,
} from './selectionContext';

export {
  getParagraphText,
  getRunText,
  getHyperlinkText,
  getTableText,
  getBodyText,
  countWords,
  countCharacters,
  getBodyWordCount,
  getBodyCharacterCount,
  getTextBefore,
  getTextAfter,
  getFormattingAtPosition,
  isPositionInHyperlink,
  getHyperlinkAtPosition,
  isHeadingStyle,
  parseHeadingLevel,
  hasImages,
  hasHyperlinks,
  hasTables,
  getParagraphs,
  getParagraphAtIndex,
  getBlockIndexForParagraph,
} from './text-utils';

export type {
  AIAction,
  AIActionRequest,
  AgentResponse,
  AgentContext,
  SelectionContext as AgentSelectionContext,
  Range,
  Position,
  ParagraphContext,
  ParagraphOutline,
  SectionInfo,
  StyleInfo,
  SuggestedAction,
  AgentCommand,
  InsertTextCommand,
  ReplaceTextCommand,
  DeleteTextCommand,
  FormatTextCommand,
  FormatParagraphCommand,
  InsertTableCommand,
  InsertImageCommand,
  InsertHyperlinkCommand,
  RemoveHyperlinkCommand,
  InsertParagraphBreakCommand,
  MergeParagraphsCommand,
  SplitParagraphCommand,
  SetVariableCommand,
  ApplyStyleCommand,
  ApplyVariablesCommand,
} from '../types/agentApi';

export {
  createCollapsedRange,
  createRange,
  isPositionInRange,
  comparePositions,
  getActionLabel,
  getActionDescription,
  createCommand,
  DEFAULT_AI_ACTIONS,
} from '../types/agentApi';

// Content-control (SDT) addressing — discover and edit block + inline controls by tag.
export {
  findContentControls,
  findContentControl,
  getContentControlText,
  setContentControlContent,
  removeContentControl,
  ContentControlNotFoundError,
  ContentControlLockedError,
  ContentControlTypeError,
  ContentControlBoundError,
  ContentControlKindError,
  type ContentControlFilter,
  type ContentControlInfo,
  type ContentControlLocation,
  type FindContentControlsOptions,
} from './contentControls';
export {
  createContentControl,
  ContentControlCreateError,
  type CreateContentControlTarget,
  type NewContentControlProps,
} from './createContentControl';
export {
  setContentControlValue,
  formatSdtDate,
  ContentControlValueError,
  type ContentControlValue,
} from './contentControlValues';
export {
  addRepeatingSectionItem,
  removeRepeatingSectionItem,
  isRepeatingSection,
  isRepeatingSectionItem,
  RepeatingSectionError,
} from './repeatingSection';
