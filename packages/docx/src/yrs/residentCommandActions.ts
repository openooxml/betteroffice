import { applyCommandBatch, commandBatchRequest } from './residentCommandBatches';
import { createProposalRegistry } from './proposals';
import { applyReviewCommand, grantedHostProposal, preflightReviewCommand, reviewCommandKeys, type ReviewCommand } from './residentReviewCommands';
import { planResidentListChanges } from './residentCommandLists';
import { breakRefusedWhileSuggesting, leadingBlockCount } from './sectionBreak';
import { tableBoundaryDeleteTarget } from './tableBoundaryDelete';
import type { YrsInlineFormatDelta, YrsLoc, YrsParagraphAttrs, YrsRunMark, YrsSelectionContext, YrsStoryRange, YrsStorySegment } from './index';
import {
  commandSelectionRange, mintCommandSelection, readCommandSelection, registerCommandAnchorState, resolveCommandSelection,
} from './residentCommandAnchors';
import { projectDocxEditState, withStoredMarks } from './residentCommandEditState';
import { resolveDocxSelectionIntent, selectResidentPartPage } from './residentCommandSelection';
import { applyCellNavigation, resolveCellNavigation } from './residentCommandCells';
import { cellsForCaret, readCellSelection, resolveLiveCell, tableRangeFor } from './residentCommandTableSelection';
import {
  applyTableCommand, isTableCommandKind, liveTableAncestors, nearTable,
  preflightTableCommand, syncCellSelectionWithCaret, type TableCommandIntent,
} from './residentCommandTables';
import { validPartTarget, resolvePart, ensurePartDocument, removePartDocument } from './residentParts';
import { validControlValue, resolveControlHandle } from './residentControls';
import { propertyCommand, validPropertyCommand, propertyCommandTarget, applyPropertyCommand } from './residentPropertyCommands';
import { validSectionMargins, nativeTabStops, residentSectionPropertiesAt } from './residentSections';
import { validFontEnvironment } from './residentFontEnvironment';
import { resolveReservations, retainReservationOwners, validReservationsPayload } from './reservations';
import { isAnchoredImagePayload, resolveResidentObject } from './residentObjectGeometry';
import { createYrsPositionProjection, type YrsPositionProjection } from './yrsPositionProjection';
import { applyClipboardFragment, validClipboardFragment } from './clipboardFragment';
import { MAX_BUFFERED_BYTES, MAX_FRAME_BYTES } from '../collaboration/protocol';
import { PIPE_OPTION_KEYS, validPipeCursor, validPipePayload } from '../../../../shared/collaboration-pipe/protocol';
import type {
  DocxCommandApplyContext,
  DocxCommandApplyResult,
  DocxCommandEnvelope,
  DocxCommandIntent,
  DocxCommandKind,
  DocxCommandPreflightResult,
  DocxCommandSelection,
  PreparedDocxCommand,
  ResidentCommandHooksFactory,
  ResidentCommandHooksEnvironment,
  ResidentDocxCommandEngine,
  ResidentDocxCommandState,
} from './residentCommandProtocol';

const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u;
const PROTECTED_MARKS = new Set(['ins', 'del', 'hyperlink']);
const IDENTITY_KEYS = new Set(['paraId', '_kind', 'ooxmlParaId', 'sourceParaId', 'paraOrigin']);

function object(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value) &&
    (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
}

function string(value: unknown): value is string {
  return typeof value === 'string' && !LONE_SURROGATE.test(value);
}

function json(value: unknown, seen = new Set<object>()): boolean {
  if (value === null || typeof value === 'boolean' || string(value)) return true;
  if (typeof value === 'number') return Number.isFinite(value);
  if (!Array.isArray(value) && !object(value)) return false;
  if (seen.has(value)) return false;
  seen.add(value);
  const valid = Object.entries(value).every(([key, entry]) => string(key) && json(entry, seen));
  seen.delete(value);
  return valid;
}

function fields(value: unknown, validators: Record<string, (value: unknown) => boolean>): boolean {
  return object(value) && Object.entries(value).every(([key, entry]) =>
    Object.hasOwn(validators, key) && validators[key]!(entry));
}

function validParagraphAnchor(value: unknown): boolean {
  if (!object(value)) return false;
  if (value.kind === 'session') return fields(value, { kind: string, sessionId: string, story: string, paraId: string }) &&
    ['sessionId', 'story', 'paraId'].every(key => string(value[key]));
  if (value.kind === 'source') return fields(value, {
    kind: string, packageSha256: string, partUri: string,
    paragraphOrdinal: entry => Number.isSafeInteger(entry) && Number(entry) >= 0 && Number(entry) <= 0xffffffff,
  }) && string(value.packageSha256) && string(value.partUri) && Number.isSafeInteger(value.paragraphOrdinal);
  if (value.kind !== 'persisted' || !fields(value, { kind: string, story: object, paraId: string }) || !string(value.paraId)) return false;
  const story = value.story;
  return fields(story, { kind: string, partUri: string, itemId: string }) && object(story) && string(story.partUri) &&
    (['body', 'header', 'footer'].includes(String(story.kind)) ||
      ['footnote', 'endnote', 'comment'].includes(String(story.kind)) && string(story.itemId));
}

const nullable = (validate: (value: unknown) => boolean) => (value: unknown) => value === null || validate(value);
const boolean = (value: unknown) => typeof value === 'boolean';
const number = (value: unknown) => typeof value === 'number' && Number.isFinite(value);
const positive = (value: unknown) => number(value) && (value as number) > 0 && Number.isFinite((value as number) * 2);
const oneOf = (...values: string[]) => (value: unknown) => typeof value === 'string' && values.includes(value);

function validDelta(value: unknown): value is YrsInlineFormatDelta {
  return fields(value, {
    bold: nullable(boolean), italic: nullable(boolean),
    underline: nullable((entry) => boolean(entry) || fields(entry, { style: string, color: string })),
    strike: nullable((entry) => boolean(entry) || fields(entry, { double: boolean })),
    color: nullable((entry) => fields(entry, { rgb: string, themeColor: string }) &&
      object(entry) && ((typeof entry.rgb === 'string') !== (typeof entry.themeColor === 'string'))),
    highlight: nullable(string), fontSize: nullable(positive),
    fontFamily: nullable((entry) => fields(entry, { ascii: string, hAnsi: string }) && object(entry) && string(entry.ascii)),
    other: (entry) => object(entry) && json(entry) && Object.keys(entry).every((key) => !PROTECTED_MARKS.has(key)),
  });
}

function validBorders(value: unknown): boolean {
  if (!object(value)) return false;
  const sides = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'];
  return Object.entries(value).every(([key, entry]) => {
    if (!sides.includes(key)) return false;
    if (entry === null) return true;
    if (!object(entry)) return false;
    return fields(entry, {
      style: string,
      size: (size) => number(size) && (size as number) >= 0,
      color: (color) => object(color) && fields(color, { rgb: string }) &&
        typeof (color as { rgb?: unknown }).rgb === 'string' && /^#?[0-9a-fA-F]{6}$/.test((color as { rgb: string }).rgb),
    }) && string((entry as { style?: unknown }).style) && (entry as { style: string }).style.trim().length > 0;
  });
}

function validTablePatch(value: unknown): boolean {
  if (!object(value)) return false;
  return fields(value, {
    width: nullable((entry) => number(entry) && (entry as number) >= 0 && Number.isInteger(entry as number)),
    widthType: nullable(oneOf('auto', 'dxa', 'pct', 'nil')),
    justification: nullable(oneOf('left', 'center', 'right')),
  });
}

function validMark(value: unknown): value is YrsRunMark {
  if (!object(value)) return false;
  if (['bold', 'italic', 'underline', 'superscript', 'subscript'].includes(String(value.type))) {
    return fields(value, { type: string });
  }
  return fields(value, { type: string, value: value.type === 'fontSize' ? positive : string }) &&
    ['fontFamily', 'fontSize', 'color'].includes(String(value.type)) && value.value !== undefined;
}

function validParagraphAttrs(value: unknown): boolean {
  if (!object(value) || !json(value) || Object.keys(value).some((key) => IDENTITY_KEYS.has(key))) return false;
  const validators: Record<string, (value: unknown) => boolean> = {
    alignment: nullable(oneOf('left', 'center', 'right', 'both', 'distribute', 'mediumKashida', 'highKashida', 'lowKashida', 'thaiDistribute')),
    lineSpacingRule: nullable(oneOf('auto', 'exact', 'atLeast')),
    hangingIndent: nullable(boolean),
    tabs: nullable((entry) => Array.isArray(entry) && entry.every((tab) => fields(tab, {
      position: number, alignment: oneOf('left', 'center', 'right', 'decimal', 'bar', 'clear', 'num'),
      leader: oneOf('none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'),
    }) && object(tab) && number(tab.position) && string(tab.alignment))),
    defaultTextFormatting: nullable((entry) => object(entry) && json(entry)),
    numPr: nullable((entry) => fields(entry, { numId: number, ilvl: number })),
    numPrFromStyle: nullable((entry) => fields(entry, { numId: number, ilvl: number })),
    listMarkerColor: nullable((entry) => object(entry) && json(entry)),
    listMarkerSuffix: nullable(oneOf('tab', 'space', 'nothing')),
    listLevelNumFmts: nullable((entry) => Array.isArray(entry) && entry.every(string)),
    other: (entry) => object(entry) && json(entry) && Object.keys(entry).every((key) => !IDENTITY_KEYS.has(key)),
  };
  for (const key of ['lineSpacing', 'spaceBefore', 'spaceAfter', 'indentLeft', 'indentRight', 'indentFirstLine',
    'listMarkerFontSize', 'listAbstractNumId', 'listStartOverride']) validators[key] = nullable(number);
  for (const key of ['bidi', 'listIsBullet', 'listMarkerHidden', 'listMarkerBold', 'listMarkerItalic']) validators[key] = nullable(boolean);
  for (const key of ['listNumFmt', 'listMarker', 'listMarkerFontFamily']) validators[key] = nullable(string);
  return fields(value, validators);
}

/** Admission policy of one command kind. */
interface DocxCommandPolicy {
  /** Payload keys besides the shared `selection` and `suggesting` keys. */
  keys: readonly string[];
  /** Refuses `selection` and `suggesting` unless `keys` lists them. */
  exact?: true;
  /** Only a collaborating session admits the kind; others refuse it as unsupported. */
  session?: 'collaboration';
  /** A remote merge: admitted on any basis and while read-only, and kept out of local undo. */
  origin?: 'remote';
  /** Whole-payload check; failing it refuses the payload as invalid. */
  payload?(payload: Record<string, unknown>): boolean;
}

/** Mirrors docx-edit's `MAX_UPDATE_BYTES`: the engine refuses a larger peer update. */
export const MAX_UPDATE_BYTES = 64 * 1024 * 1024;
/** What one chunked sync hands over: updates held within the largest `maxSyncBytes`, then its SyncStep2. */
export const MAX_SYNC_CHUNKS_BYTES = MAX_BUFFERED_BYTES + MAX_UPDATE_BYTES;

const peerUpdate = (update: unknown): update is ArrayBuffer =>
  update instanceof ArrayBuffer && update.byteLength > 0 && update.byteLength <= MAX_UPDATE_BYTES;

const PIPE_SHAPE = { validCursor: validPipeCursor, optionKeys: PIPE_OPTION_KEYS };

const DOCX_COMMAND_POLICIES = {
  applyRemoteUpdate: { keys: ['update'], exact: true, session: 'collaboration', origin: 'remote',
    payload: ({ update }) => peerUpdate(update) },
  applyRemoteUpdateChunks: { keys: ['updates'], exact: true, session: 'collaboration', origin: 'remote',
    payload: ({ updates }) => {
      if (!Array.isArray(updates) || updates.length === 0) return false;
      let total = 0;
      for (const update of updates) {
        if (!peerUpdate(update)) return false;
        total += update.byteLength;
      }
      return total <= MAX_SYNC_CHUNKS_BYTES;
    } },
  collabPipe: { keys: ['provider', 'ops'], exact: true, session: 'collaboration', origin: 'remote',
    payload: (payload) => validPipePayload(payload, MAX_FRAME_BYTES, PIPE_SHAPE) },
  applyEdits: { keys: ['request', 'suggesting'], exact: true },
  setFontEnvironment: { keys: ['faces', 'chains', 'families'], exact: true },
  setContentControlValue: { keys: [] }, addRepeatingItem: { keys: [] }, removeRepeatingItem: { keys: [] },
  openPart: { keys: [] }, ensurePart: { keys: [] }, removePart: { keys: [] },
  setHyperlink: { keys: [] }, setObjectProperties: { keys: [] }, setSectionProperties: { keys: [] },
  setNoteProperties: { keys: [] }, setWatermark: { keys: [] },
  setSectionMargins: { keys: ['section', 'margins', 'expectVersion'] },
  setTabStops: { keys: ['tabs', 'expectVersion'] }, removeTabStop: { keys: ['positionTwips', 'expectVersion'] },
  formatSelection: { keys: ['delta'] }, toggleStrike: { keys: [] }, toggleList: { keys: ['kind'] },
  changeIndent: { keys: ['direction', 'preserveNumbering'] },
  insertImage: { keys: ['image'] }, insertTable: { keys: ['rows', 'columns'] }, pasteFragment: { keys: ['fragment'] },
  insertPageBreak: { keys: [] }, insertSectionBreak: { keys: ['type'] },
  insertTableRow: { keys: ['side'] }, insertTableColumn: { keys: ['side'] },
  deleteTableRow: { keys: [] }, deleteTableColumn: { keys: [] }, deleteTable: { keys: [] },
  mergeTableCells: { keys: [] }, splitTableCell: { keys: ['rows', 'columns'] }, setTableCellShading: { keys: ['color'] },
  setTableBorders: { keys: ['borders'] }, setTableProperties: { keys: ['op', 'patch', 'expectVersion'] },
  resizeObject: { keys: ['handle', 'widthEmu', 'heightEmu', 'widthTwips', 'hOffsetEmu', 'vOffsetEmu', 'expectVersion'], exact: true },
  moveObject: { keys: ['handle', 'hOffsetEmu', 'vOffsetEmu', 'expectVersion'], exact: true },
  undo: { keys: [], exact: true }, redo: { keys: [], exact: true },
  replaceMatches: { keys: ['targets', 'text', 'expectVersion', 'suggesting'], exact: true },
  addComment: { keys: reviewCommandKeys.addComment, exact: true },
  updateComment: { keys: reviewCommandKeys.updateComment, exact: true },
  setCommentRanges: { keys: reviewCommandKeys.setCommentRanges, exact: true },
  removeComment: { keys: reviewCommandKeys.removeComment, exact: true },
  acceptRevision: { keys: reviewCommandKeys.acceptRevision, exact: true },
  rejectRevision: { keys: reviewCommandKeys.rejectRevision, exact: true },
  propose: { keys: reviewCommandKeys.propose, exact: true },
  setProposalStates: { keys: reviewCommandKeys.setProposalStates, exact: true },
  withdrawProposals: { keys: reviewCommandKeys.withdrawProposals, exact: true },
  select: { keys: ['selection', 'object'] }, insertText: { keys: ['text'] }, replaceSelection: { keys: ['text'] },
  navigate: { keys: ['unit', 'direction', 'extend', 'goalX', 'distance'], exact: true },
  navigateCell: { keys: ['direction', 'extend', 'createLastRow'] },
  selectTableCells: { keys: ['target'] },
  selectPoint: { keys: ['page', 'x', 'y', 'layoutRevision', 'frameEpoch', 'clickCount', 'modifiers', 'story', 'cellId'], exact: true },
  extendSelection: { keys: ['page', 'x', 'y', 'layoutRevision', 'frameEpoch', 'clickCount', 'modifiers', 'story', 'cellId'], exact: true },
  deleteBackward: { keys: ['count'] }, deleteForward: { keys: ['count'] }, paragraphBreak: { keys: [] }, lineBreak: { keys: [] },
  toggleMark: { keys: ['mark'] }, formatRange: { keys: ['delta'] }, clearFormatting: { keys: [] },
  setParagraphAttrs: { keys: ['attrs'] }, applyParagraphStyle: { keys: ['styleId'] },
  setReadOnly: { keys: ['readOnly'], exact: true },
  setRenderEnvironment: { keys: ['themeColors', 'showHiddenText', 'showTableGridlines', 'reservations'], exact: true },
} satisfies Record<DocxCommandKind, DocxCommandPolicy>;

function commandPolicy(kind: string): DocxCommandPolicy | undefined {
  return Object.hasOwn(DOCX_COMMAND_POLICIES, kind) ? DOCX_COMMAND_POLICIES[kind as DocxCommandKind] : undefined;
}

function validCommandIntent(
  command: DocxCommandEnvelope, error: { reason: string }, environment?: ResidentCommandHooksEnvironment
): command is DocxCommandIntent {
  const refuse = (reason: string): false => { error.reason = reason; return false; };
  const payload: unknown = command.payload;
  if (!object(payload)) return refuse('invalid-payload');
  const policy = commandPolicy(command.kind);
  if (policy?.session === 'collaboration' && !environment?.collaboration) return refuse('unsupported-command');
  if (['setContentControlValue', 'addRepeatingItem', 'removeRepeatingItem'].includes(command.kind)) {
    const keys = command.kind === 'setContentControlValue' ? ['handle', 'value', 'expectVersion'] : ['handle', 'expectVersion'];
    return Object.keys(payload).every(key => keys.includes(key)) && string(payload.handle) && string(payload.expectVersion) &&
      (command.kind !== 'setContentControlValue' || validControlValue(payload.value)) || refuse('invalid-control-value');
  }
  if (['openPart', 'ensurePart', 'removePart'].includes(command.kind)) return validPartTarget(payload) || refuse('invalid-part-target');
  if (propertyCommand(command.kind)) return validPropertyCommand(command) || refuse('invalid-properties');
  if (!policy || Object.keys(payload).some((key) => !policy.keys.includes(key) &&
    (policy.exact || !['selection', 'suggesting'].includes(key)))) return refuse('invalid-payload');
  if (policy.payload && !policy.payload(payload)) return refuse('invalid-payload');
  if (payload.suggesting !== undefined && (!fields(payload.suggesting, { name: string, date: string }) ||
    !object(payload.suggesting) || !string(payload.suggesting.name) || !string(payload.suggesting.date))) return refuse('invalid-author');
  if (payload.hostProposal !== undefined && !boolean(payload.hostProposal)) return refuse('invalid-payload');
  switch (command.kind) {
    case 'applyEdits':
      if (!object(payload.request) || !string(payload.request.expectVersion) || !Array.isArray(payload.request.steps) || !json(payload.request)) return refuse('invalid-payload');
      break;
    case 'setFontEnvironment':
      if (!validFontEnvironment(payload)) return refuse('invalid-font-environment');
      break;
    case 'setSectionMargins':
      if (!string(payload.section) || !string(payload.expectVersion) || !validSectionMargins(payload.margins)) return refuse('invalid-margins');
      break;
    case 'setTabStops':
      if (!string(payload.expectVersion) || payload.tabs === undefined || !validParagraphAttrs({ tabs: payload.tabs })) return refuse('invalid-tab-stops');
      break;
    case 'removeTabStop':
      if (!string(payload.expectVersion) || !number(payload.positionTwips) || Number(payload.positionTwips) < 0) return refuse('invalid-tab-stops');
      break;
    case 'resizeObject': case 'moveObject':
      if (!string(payload.handle) || !string(payload.expectVersion) || (command.kind === 'resizeObject'
        ? payload.widthTwips !== undefined
          ? !positive(payload.widthTwips) || payload.widthEmu !== undefined || payload.heightEmu !== undefined || payload.hOffsetEmu !== undefined || payload.vOffsetEmu !== undefined
          : !positive(payload.widthEmu) || !positive(payload.heightEmu) ||
            ((payload.hOffsetEmu !== undefined || payload.vOffsetEmu !== undefined) && (!number(payload.hOffsetEmu) || !number(payload.vOffsetEmu)))
        : !number(payload.hOffsetEmu) || !number(payload.vOffsetEmu))) return refuse('invalid-payload');
      break;
    case 'replaceMatches':
      if (!string(payload.text) || !string(payload.expectVersion) || !Array.isArray(payload.targets) || !json(payload.targets) || !payload.targets.every(target => object(target) && string(target.story) && oneOf('accepted', 'original')(target.view) && ['start', 'end'].every(key => object(target[key]) && string(target[key].paraId) && Number.isSafeInteger(target[key].offset) && Number(target[key].offset) >= 0))) return refuse('invalid-payload');
      break;
    case 'addComment':
      if (!string(payload.author) || !string(payload.date) || !json(payload.body)) return refuse('invalid-payload');
      break;
    case 'updateComment':
      if (!string(payload.id) || !fields(payload.patch, { body: json, author: string, date: string, done: boolean, parentId: nullable(string) })) return refuse('invalid-payload');
      break;
    case 'propose':
      if (!Array.isArray(payload.proposals) || !json(payload.proposals) || payload.proposals.some(entry => !object(entry) || !string(entry.id) || !validParagraphAnchor(entry.paragraph) || !object(entry.suggest) || !string(entry.suggest.author) || !string(entry.suggest.date) || !['insertText', 'replaceText'].includes(String(entry.op)) || (entry.op === 'insertText' ? !string(entry.text) : !string(entry.search) || !string(entry.replaceWith)))) return refuse('invalid-payload');
      break;
    case 'setProposalStates':
      if (!Array.isArray(payload.changes) || !payload.changes.every(entry => fields(entry, { id: string, state: oneOf('proposed', 'accepted', 'rejected') }) && object(entry) && string(entry.id) && string(entry.state)) || !Number.isSafeInteger(payload.expectPreviewVersion)) return refuse('invalid-payload');
      break;
    case 'withdrawProposals':
      if (!Array.isArray(payload.ids) || !payload.ids.every(string)) return refuse('invalid-payload');
      break;
    case 'removeComment': case 'setCommentRanges':
      if (!string(payload.id)) return refuse('invalid-payload');
      break;
    case 'acceptRevision': case 'rejectRevision':
      if (!string(payload.revisionId)) return refuse('invalid-payload');
      break;
    case 'navigateCell':
      if (!oneOf('backward', 'forward')(payload.direction) || !boolean(payload.extend) ||
        !boolean(payload.createLastRow)) return refuse('invalid-payload');
      break;
    case 'selectTableCells':
      if (!oneOf('table', 'row', 'column')(payload.target)) return refuse('invalid-payload');
      break;
    case 'navigate':
      if (!oneOf('character', 'word', 'line', 'page', 'paragraph', 'document-boundary')(payload.unit) ||
        !boolean(payload.extend) || !oneOf('backward', 'forward', 'up', 'down')(payload.direction) ||
        (['up', 'down'].includes(String(payload.direction)) && (!oneOf('line', 'page')(payload.unit) || !number(payload.goalX))) ||
        (payload.unit !== 'page' && payload.distance !== undefined) ||
        (payload.unit === 'page' && (!oneOf('up', 'down')(payload.direction) || !number(payload.distance) || Number(payload.distance) <= 0))) {
        return refuse('invalid-payload');
      }
      break;
    case 'selectPoint': case 'extendSelection': {
      const modifiers = payload.modifiers;
      if (!fields(payload.page, { pageId: string, pageIndex: number }) || !object(payload.page) ||
        !string(payload.page.pageId) || !number(payload.page.pageIndex) ||
        !number(payload.x) || !number(payload.y) || !number(payload.layoutRevision) || !number(payload.frameEpoch) ||
        !number(payload.clickCount) || !fields(modifiers, { shift: boolean, alt: boolean, ctrl: boolean, meta: boolean }) ||
        !object(modifiers) || !['shift', 'alt', 'ctrl', 'meta'].every((key) => boolean(modifiers[key])) ||
        (payload.story !== undefined && !string(payload.story)) || (payload.cellId !== undefined && !string(payload.cellId))) {
        return refuse('invalid-payload');
      }
      break;
    }
    case 'insertText': case 'replaceSelection':
      if (!string(payload.text) || /[\n\r\u000B\u2028\u2029]/u.test(payload.text)) return refuse('invalid-text');
      break;
    case 'deleteBackward': case 'deleteForward':
      if (payload.count !== undefined && (typeof payload.count !== 'number' ||
        !Number.isSafeInteger(payload.count) || payload.count <= 0 || payload.count > 0xffffffff)) return refuse('invalid-count');
      break;
    case 'toggleMark':
      if (!validMark(payload.mark)) return refuse('invalid-formatting');
      break;
    case 'formatSelection': case 'formatRange':
      if (!validDelta(payload.delta)) return refuse('invalid-formatting');
      break;
    case 'setParagraphAttrs':
      if (!validParagraphAttrs(payload.attrs)) return refuse('invalid-paragraph-attrs');
      break;
    case 'applyParagraphStyle':
      if (!string(payload.styleId) || !payload.styleId) return refuse('invalid-style');
      break;
    case 'toggleList':
      if (!oneOf('bullet', 'numbered')(payload.kind)) return refuse('invalid-payload');
      break;
    case 'changeIndent':
      if (!oneOf('increase', 'decrease')(payload.direction) ||
        (payload.preserveNumbering !== undefined && !boolean(payload.preserveNumbering))) return refuse('invalid-payload');
      break;
    case 'insertTable':
      if (![payload.rows, payload.columns].every(entry => typeof entry === 'number' && Number.isSafeInteger(entry) && entry > 0 && entry <= 1000) || Number(payload.rows) * Number(payload.columns) > 10_000) return refuse('invalid-table-size');
      break;
    case 'pasteFragment':
      if (!validClipboardFragment(payload.fragment)) return refuse('invalid-payload');
      break;
    case 'insertTableRow':
      if (!oneOf('above', 'below')(payload.side)) return refuse('invalid-payload');
      break;
    case 'insertTableColumn':
      if (!oneOf('left', 'right')(payload.side)) return refuse('invalid-payload');
      break;
    case 'splitTableCell':
      if (![payload.rows, payload.columns].every(entry => typeof entry === 'number' && Number.isSafeInteger(entry) && entry >= 1 && entry <= 1000) || Number(payload.rows) * Number(payload.columns) > 10_000) return refuse('invalid-table-size');
      break;
    case 'setTableCellShading':
      if (payload.color !== null && (typeof payload.color !== 'string' || !/^[0-9a-fA-F]{6}$/.test(payload.color))) return refuse('invalid-payload');
      break;
    case 'setTableBorders':
      if (!validBorders(payload.borders)) return refuse('invalid-payload');
      break;
    case 'setTableProperties':
      if (payload.op !== 'patch' || !validTablePatch(payload.patch) ||
        (payload.expectVersion !== undefined && !string(payload.expectVersion))) return refuse('invalid-payload');
      break;
    case 'insertImage':
      if (!object(payload.image) || !json(payload.image) || Object.hasOwn(payload.image, '_kind') || ['widthEmu', 'heightEmu'].some(key => payload.image && object(payload.image) && payload.image[key] !== undefined && !positive(payload.image[key]))) return refuse('invalid-image');
      break;
    case 'insertSectionBreak':
      if (!oneOf('nextPage', 'continuous', 'oddPage', 'evenPage')(payload.type)) return refuse('invalid-payload');
      break;
    case 'setReadOnly':
      if (typeof payload.readOnly !== 'boolean') return refuse('invalid-payload');
      break;
    case 'setRenderEnvironment':
      if (!object(payload.themeColors) || !Object.values(payload.themeColors).every(value =>
        typeof value === 'string' && /^[0-9a-f]{6}$/i.test(value)) || !boolean(payload.showHiddenText) ||
        (payload.showTableGridlines !== undefined && !boolean(payload.showTableGridlines)) ||
        (payload.reservations !== undefined && !validReservationsPayload(payload.reservations))) return refuse('invalid-payload');
      break;
  }
  return true;
}

function collapsed(range: YrsStoryRange): boolean {
  return range.start.paraId === range.end.paraId && range.start.offset === range.end.offset;
}

function deleteRange(
  engine: ResidentDocxCommandEngine, range: YrsStoryRange, backward: boolean, count: number, skipDeleted: boolean,
  paragraphSegments: ResidentDocxCommandEngine['paragraphSegments']
): YrsStoryRange | null {
  if (!collapsed(range)) return range;
  const spans = engine.geometryReader.paragraphSpans(range.story);
  let index = spans.findIndex(({ paraId }) => paraId === range.start.paraId);
  if (index < 0) return null;
  let offset = range.start.offset;
  const paragraphs = new Map<number, { text: string; deleted: Array<[number, number]> }>();
  const paragraphAt = (index: number) => {
    let paragraph = paragraphs.get(index);
    if (paragraph === undefined) {
      paragraph = { text: '', deleted: [] };
      for (const segment of paragraphSegments(range.story, [index])[0] ?? []) {
        const text = segment.kind === 'text' ? segment.text : segment.kind === 'pilcrow' ? '\n' : '\uFFFC';
        if (skipDeleted && !(segment.kind === 'embed' && segment.embedKind === 'table') && (segment.attributes.del != null ||
          (segment.kind === 'pilcrow' && segment.properties.pPrDel != null))) {
          paragraph.deleted.push([paragraph.text.length, paragraph.text.length + text.length]);
        }
        paragraph.text += text;
      }
      paragraphs.set(index, paragraph);
    }
    return paragraph;
  };
  const tableBoundary = () => tableBoundaryDeleteTarget({
    paragraphSpans: story => story === range.story ? spans : engine.geometryReader.paragraphSpans(story),
    paragraphSegments,
  }, { story: range.story, paraId: spans[index]!.paraId, offset }, backward ? 'backward' : 'forward');
  for (let left = count; left > 0; left -= 1) {
    const text = paragraphAt(index).text;
    const boundary = backward ? offset === 0 || text[offset - 1] === '\uFFFC' :
      offset === spans[index]!.length || text[offset] === '\uFFFC';
    if (boundary && tableBoundary()) break;
    if (backward) {
      if (offset === 0) {
        if (index === 0) break;
        offset = spans[--index]!.length;
      } else {
        const { text } = paragraphAt(index);
        offset -= 1;
        if (offset > 0 && /[\uDC00-\uDFFF]/u.test(text[offset]!) && /[\uD800-\uDBFF]/u.test(text[offset - 1]!)) offset -= 1;
      }
    } else {
      if (skipDeleted) {
        for (;;) {
          const deleted = paragraphAt(index).deleted.find(([start, end]) => start <= offset && offset < end);
          if (!deleted) break;
          offset = deleted[1];
          if (offset > spans[index]!.length) {
            if (index === spans.length - 1) {
              offset = spans[index]!.length;
              break;
            }
            index += 1;
            offset = 0;
          }
        }
        if (tableBoundary()) break;
      }
      if (offset === spans[index]!.length) {
        if (index === spans.length - 1) break;
        index += 1;
        offset = 0;
      } else {
        offset += paragraphAt(index).text.codePointAt(offset)! > 0xffff ? 2 : 1;
      }
    }
  }
  const target = { paraId: spans[index]!.paraId, offset };
  const expanded = {
    story: range.story, start: backward ? target : range.start, end: backward ? range.end : target,
  };
  return commandSelectionRange(engine, {
    anchor: { story: range.story, ...expanded.start }, focus: { story: range.story, ...expanded.end },
  }) ? expanded : null;
}

export function preflightDocxCommand(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  command: DocxCommandEnvelope,
  context: DocxCommandApplyContext,
  environment?: ResidentCommandHooksEnvironment
): DocxCommandPreflightResult {
  const refuse = (reason: string): DocxCommandPreflightResult => {
    const corrected = state.selection ? readCommandSelection(engine, state) : null;
    return { status: 'refused', reason, selection: corrected?.ok ? corrected.selection : null,
      editState: projectDocxEditState(engine, state, corrected) };
  };
  const readOnly = state.readOnly && !context.replayPreviouslyApplied;
  const error = { reason: 'invalid-payload' };
  if (!validCommandIntent(command, error, environment)) return refuse(error.reason);
  if (commandPolicy(command.kind)?.origin === 'remote') {
    return { status: 'ready', prepared: { command, selection: state.selection, range: null } };
  }
  if (command.kind === 'setRenderEnvironment') {
    if (!environment?.setRenderEnvironment) return refuse('unsupported-command');
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  if (command.kind === 'setFontEnvironment') {
    if (!environment?.supportsFontEnvironment) return refuse('unsupported-command');
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  if (command.kind === 'setContentControlValue' || command.kind === 'addRepeatingItem' || command.kind === 'removeRepeatingItem') {
    if (readOnly) return refuse('read-only');
    if (command.payload.expectVersion !== engine.proposalEngine.version()) return refuse('stale-version');
    const target = resolveControlHandle(engine, command.payload.handle);
    if (!target) return refuse('stale-control');
    if (target.control.effectiveLock.content || !target.control.effectiveLock.known || target.control.dataBound) return refuse('locked-target');
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  if (command.kind === 'openPart' || command.kind === 'ensurePart' || command.kind === 'removePart') {
    if (readOnly && command.kind !== 'openPart') return refuse('read-only');
    if (command.payload.expectVersion !== engine.proposalEngine.version()) return refuse('stale-version');
    if (!resolvePart(engine, command.payload) && (command.kind !== 'ensurePart' || !ensurePartDocument(engine, command.payload))) return refuse('unknown-part');
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  if (propertyCommand(command.kind)) {
    if (readOnly) return refuse('read-only');
    if (command.basis.epoch !== context.basis.epoch) return refuse('stale-epoch');
    if (!('expectVersion' in command.payload) || command.payload.expectVersion !== engine.proposalEngine.version()) return refuse('stale-version');
    const reason = propertyCommandTarget(engine, command);
    if (reason) return refuse(reason);
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  if (command.kind === 'setSectionMargins' || command.kind === 'setTabStops' || command.kind === 'removeTabStop') {
    if (readOnly) return refuse('read-only');
    if ('suggesting' in command.payload && command.payload.suggesting !== undefined) return refuse('unsupported-suggesting');
    if (command.payload.expectVersion !== engine.proposalEngine.version()) return refuse('stale-version');
    if (command.kind === 'setSectionMargins') {
      if (!engine.sectionProperties().sections.some(section => section.section === command.payload.section)) return refuse('unknown-section');
      return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
    }
  }
  switch (command.kind) {
    case 'navigate': case 'selectPoint': case 'extendSelection': {
      if (!command.basis || !Number.isSafeInteger(command.basis.epoch) || command.basis.epoch < 0 ||
        !Number.isSafeInteger(command.basis.version) || command.basis.version < 0 ||
        command.basis.epoch !== context.basis.epoch || command.basis.version !== context.basis.version ||
        command.basis.version > context.currentVersion) return refuse('stale-anchor');
      if (command.kind === 'navigate') {
        const base = resolveCommandSelection(engine, state, command, context);
        if (!base.ok) return refuse(base.reason);
      }
      const resolved = resolveDocxSelectionIntent({
        engine, state, command, context, readLayout: environment?.readSelectionLayout,
      });
      return { status: 'ready', prepared: {
        command: structuredClone(command), selection: resolved.selection, range: resolved.range,
        cells: resolved.cells ?? cellsForCaret(engine, resolved.selection.focus.story),
      } };
    }
    case 'selectTableCells': {
      const resolved = resolveCommandSelection(engine, state, command, context);
      if (!resolved.ok) return refuse(resolved.reason);
      const focused = resolveLiveCell(engine, resolved.selection.focus.story);
      const cells = focused ? tableRangeFor(engine, focused, command.payload.target) : null;
      if (!focused || !cells) return refuse('invalid-cell-target');
      const at = resolved.selection.focus;
      const selection = { anchor: at, focus: at };
      const range = commandSelectionRange(engine, selection);
      if (!range) return refuse('invalid-selection');
      return { status: 'ready', prepared: { command: structuredClone(command), selection, range, cells } };
    }
  }
  if (command.kind === 'applyEdits') {
    if (readOnly) return refuse('read-only');
    const validation = engine.proposalEngine.validateEdits(commandBatchRequest(command.payload));
    if (!validation.ok) return refuse(validation.failure.code);
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  if (command.kind === 'replaceMatches') {
    if (readOnly) return refuse('read-only');
    if (command.payload.expectVersion !== engine.proposalEngine.version()) return refuse('stale-version');
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  if (command.kind === 'resizeObject' || command.kind === 'moveObject') {
    if (readOnly) return refuse('read-only');
    if (command.basis.epoch !== context.basis.epoch) return refuse('stale-epoch');
    if (command.payload.expectVersion !== engine.proposalEngine.version()) return refuse('stale-version');
    const target = resolveResidentObject(engine, command.payload.handle);
    if (!target) return refuse('stale-object');
    if (target.kind === 'cell' && command.kind === 'resizeObject' && 'widthTwips' in command.payload) {
      if (!target.cell) return refuse('stale-object');
    } else if (target.kind !== 'image' || (command.kind === 'resizeObject' && 'widthTwips' in command.payload)) return refuse('unsupported-object');
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  if (command.kind === 'undo' || command.kind === 'redo') {
    if (readOnly) return refuse('read-only');
    if (command.basis.epoch !== context.basis.epoch) return refuse('stale-epoch');
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  if (Object.hasOwn(reviewCommandKeys, command.kind)) {
    if (readOnly && !grantedHostProposal(command as ReviewCommand)) return refuse('read-only');
    const reason = preflightReviewCommand(engine, state, command as ReviewCommand);
    if (reason) return refuse(reason);
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  if (command.kind === 'setReadOnly') {
    if (!command.basis || !Number.isSafeInteger(command.basis.epoch) || command.basis.epoch < 0 ||
      command.basis.epoch !== context.basis.epoch || command.basis.version !== context.basis.version ||
      !Number.isSafeInteger(command.basis.version) || command.basis.version < 0 ||
      command.basis.version > context.currentVersion) return refuse('stale-anchor');
    if (command.expectVersion && command.basis.version !== context.currentVersion) return refuse('stale-version');
    return { status: 'ready', prepared: { command: structuredClone(command), selection: state.selection, range: null } };
  }
  let objectSelection: DocxCommandSelection | undefined;
  if (command.kind === 'select' && command.payload.object !== undefined) {
    // Select a picture by its issued handle: the handle's sticky position finds the embed, and a stale version is refused.
    const picked = command.payload.object as unknown;
    if (command.payload.selection !== undefined || !fields(picked, { handle: string, expectVersion: string }) || !object(picked) ||
      !string(picked.handle) || !string(picked.expectVersion)) return refuse('invalid-payload');
    if (picked.expectVersion !== engine.proposalEngine.version()) return refuse('stale-version');
    const target = resolveResidentObject(engine, picked.handle);
    if (!target || target.kind !== 'image') return refuse('stale-object');
    const at = { story: target.story, paraId: target.paraId, offset: target.offset };
    objectSelection = { anchor: at, focus: { ...at, offset: at.offset + 1 } };
  }
  if ((command.kind === 'select' || command.kind === 'formatRange') && command.payload.selection === undefined && !objectSelection) return refuse('invalid-selection');
  if (isTableCommandKind(command.kind)) {
    if (readOnly) return refuse('read-only');
    if (command.kind === 'setTableProperties' && command.payload.expectVersion !== undefined &&
      command.payload.expectVersion !== engine.proposalEngine.version()) return refuse('stale-version');
    const tableResolved = resolveCommandSelection(engine, state, command, context);
    if (!tableResolved.ok) return refuse(tableResolved.reason);
    const tabled = preflightTableCommand(engine, command as TableCommandIntent, tableResolved.selection);
    if ('reason' in tabled) return refuse(tabled.reason);
    return { status: 'ready', prepared: {
      command: structuredClone(command), selection: tableResolved.selection, range: tableResolved.range, table: tabled.target,
    } };
  }
  const resolved = resolveCommandSelection(engine, state,
    objectSelection ? { ...command, payload: { selection: objectSelection } } as DocxCommandEnvelope : command, context);
  if (!resolved.ok) return refuse(resolved.reason);
  if (command.kind === 'navigateCell') {
    const cellNavigation = resolveCellNavigation(engine, resolved.selection, command.payload, readOnly);
    if (cellNavigation === 'read-only') return refuse('read-only');
    if (!cellNavigation) return refuse('invalid-cell-target');
    return { status: 'ready', prepared: {
      command: structuredClone(command), selection: resolved.selection, range: resolved.range, cellNavigation,
    } };
  }
  if (command.kind === 'select') {
    return { status: 'ready', prepared: {
      command: structuredClone(objectSelection ? { ...command, payload: { selection: objectSelection } } : command), selection: resolved.selection, range: resolved.range,
      cells: cellsForCaret(engine, resolved.selection.focus.story),
    } };
  }
  if (readOnly) return refuse('read-only');
  if (breakRefusedWhileSuggesting(command.kind, 'suggesting' in command.payload && command.payload.suggesting !== undefined)) return refuse('unsupported-suggesting');
  if (command.kind === 'insertPageBreak' || command.kind === 'insertSectionBreak') {
    if (resolved.range.story !== 'body') return refuse('invalid-selection');
    const index = engine.geometryReader.paragraphSpans('body').findIndex(span => span.paraId === resolved.range.start.paraId);
    const leading = leadingBlockCount(engine.paragraphSegments('body', [index])[0] ?? []);
    if (resolved.range.start.offset < leading) return refuse('invalid-selection');
    if (command.kind === 'insertSectionBreak') {
      const properties = residentSectionPropertiesAt(engine, { story: 'body', ...resolved.range.start });
      if (!properties) return refuse('invalid-selection');
      return { status: 'ready', prepared: { command: structuredClone(command), selection: resolved.selection, range: resolved.range,
        sectionBreakProperties: properties } };
    }
  }
  if (command.kind === 'toggleList' || command.kind === 'changeIndent') return { status: 'ready', prepared: {
    command: structuredClone(command), selection: resolved.selection, range: resolved.range,
    paragraphChanges: planResidentListChanges(engine, resolved.range, command.payload),
  } };
  let range = resolved.range;
  if ((command.kind === 'insertText' || command.kind === 'replaceSelection' || command.kind === 'lineBreak') && state.storedMarks &&
    (typeof state.storedMarks.clear !== 'boolean' || !validDelta(state.storedMarks.delta))) return refuse('invalid-stored-marks');
  if (command.kind === 'deleteBackward' || command.kind === 'deleteForward') {
    const story = range.story;
    const segments = new Map<number, YrsStorySegment[]>();
    const paragraphSegments: ResidentDocxCommandEngine['paragraphSegments'] = (readStory, indices) => {
      if (readStory !== story) return engine.paragraphSegments(readStory, indices);
      return indices.map(index => {
        let paragraph = segments.get(index);
        if (paragraph === undefined) {
          paragraph = engine.paragraphSegments(story, [index])[0] ?? [];
          segments.set(index, paragraph);
        }
        return paragraph;
      });
    };
    if (collapsed(range)) {
      const target = tableBoundaryDeleteTarget({
        paragraphSpans: story => engine.geometryReader.paragraphSpans(story),
        paragraphSegments,
      }, { story: range.story, ...range.start }, command.kind === 'deleteBackward' ? 'backward' : 'forward');
      if (target) return { status: 'ready', prepared: {
        command: structuredClone(command), selection: { anchor: target, focus: target },
        range: {
          story: target.story, start: { paraId: target.paraId, offset: target.offset },
          end: { paraId: target.paraId, offset: target.offset },
        },
      } };
    }
    const expanded = deleteRange(engine, range, command.kind === 'deleteBackward', command.payload.count ?? 1,
      command.kind === 'deleteForward' && command.payload.suggesting !== undefined, paragraphSegments);
    if (!expanded) return refuse('unsupported-delete-target');
    range = expanded;
  }
  return { status: 'ready', prepared: {
    command: structuredClone(command), selection: resolved.selection, range,
  } };
}

function sameSelection(left: DocxCommandSelection | null, right: DocxCommandSelection): boolean {
  return !!left && (['anchor', 'focus'] as const).every((key) => left[key].story === right[key].story &&
    left[key].paraId === right[key].paraId && left[key].offset === right[key].offset);
}

const CELL_RESET_KINDS = new Set(['insertText', 'replaceSelection', 'deleteBackward', 'deleteForward', 'paragraphBreak', 'lineBreak']);

function resetCellSelectionForEdit(engine: ResidentDocxCommandEngine, state: ResidentDocxCommandState, kind: string): void {
  if (state.selection && CELL_RESET_KINDS.has(kind)) writeCellSelection(engine, cellsForCaret(engine, state.selection.focus.story));
}

function writeCellSelection(engine: ResidentDocxCommandEngine, cells: PreparedDocxCommand['cells']): void {
  try {
    if (cells) engine.setCellSelection(cells);
    else engine.clearCellSelection();
  } catch (error) {
    if (error instanceof WebAssembly.RuntimeError) throw error;
  }
}

function pastLeadingBlocks(engine: ResidentDocxCommandEngine, at: YrsLoc): YrsLoc | null {
  if (!engine.geometryReader.hasStory(at.story)) return null;
  const spans = engine.geometryReader.paragraphSpans(at.story);
  const index = spans.findIndex(span => span.paraId === at.paraId);
  if (index < 0) return null;
  const leading = leadingBlockCount(engine.paragraphSegments(at.story, [index])[0] ?? []);
  return { ...at, offset: Math.min(spans[index]!.length, Math.max(leading, at.offset)) };
}

export function repairCommandSelection(engine: ResidentDocxCommandEngine, state: ResidentDocxCommandState): boolean {
  const cells = readCellSelection(engine);
  const preferred = engine.clearCellSelection();
  const projections = new Map<string, YrsPositionProjection | null>();
  const projection = (story: string) => {
    const root = story === 'body' || story.startsWith('body:') ? 'body' : story.split(':').slice(0, 2).join(':');
    if (!projections.has(root)) projections.set(root, createYrsPositionProjection(engine.geometryReader, root));
    return projections.get(root);
  };
  const safe = (at: YrsLoc): YrsLoc | null => {
    const clamped = pastLeadingBlocks(engine, at);
    return clamped && projection(at.story)?.positionForLoc(at) != null ? clamped : null;
  };
  const resolved = state.selection ? readCommandSelection(engine, state) : null;
  const collapsed = resolved?.ok && sameSelection({ anchor: resolved.selection.anchor, focus: resolved.selection.anchor }, resolved.selection);
  let selected = resolved?.ok && [resolved.selection.anchor, resolved.selection.focus].every(at => {
    const valid = safe(at);
    return valid && (!collapsed || valid.offset === at.offset);
  }) ? resolved.selection : null;
  if (!selected) {
    const focus = state.selection?.focus;
    const stories: string[] = [];
    for (let story = focus?.story; story; story = story.includes(':') ? story.slice(0, story.lastIndexOf(':')) : undefined) stories.push(story);
    if (!stories.includes('body')) stories.push('body');
    for (const story of stories) {
      if (!engine.geometryReader.hasStory(story)) continue;
      let at = preferred?.story === story ? safe(preferred) : focus?.story === story ? safe(focus) : null;
      for (const paragraph of at ? [] : engine.geometryReader.paragraphSpans(story)) {
        at = safe({ story, paraId: paragraph.paraId, offset: 0 });
        if (at) break;
      }
      if (at) { selected = { anchor: at, focus: at }; break; }
    }
  }
  if (!selected) return false;
  const changed = !sameSelection(state.selection, selected);
  engine.setSelection(selected.anchor, selected.focus);
  state.selection = mintCommandSelection(engine, state, selected);
  const tables = projection(selected.focus.story)?.allTables() ?? [];
  for (const table of tables) {
    const cell = table.cells.find(cell => cell.story === selected!.focus.story);
    if (!cell) continue;
    const at = { story: table.story, tableIndex: table.tableIndex, row: cell.row, column: cell.column };
    const head = cells?.head;
    const keep = head?.story === at.story && head.tableIndex === at.tableIndex && head.row === at.row && head.column === at.column;
    engine.setCellSelection(keep ? cells! : { anchor: at, head: at });
    break;
  }
  if (changed) state.storedMarks = null;
  if (state.activePart && !engine.geometryReader.hasStory(state.activePart.story)) state.activePart = null;
  return true;
}

function paragraphAttrs(attrs: YrsParagraphAttrs): YrsParagraphAttrs {
  if (typeof attrs.hangingIndent !== 'boolean') return attrs;
  const { hangingIndent, ...rest } = attrs;
  return { ...rest, other: { hangingIndent, ...attrs.other } };
}

function markDelta(mark: YrsRunMark, context: NonNullable<ReturnType<typeof projectDocxEditState>['toolbar']>): YrsInlineFormatDelta {
  switch (mark.type) {
    case 'bold': case 'italic': case 'underline': return { [mark.type]: context[mark.type] !== true };
    case 'superscript': return { other: { superscript: context.superscript !== true, ...(context.superscript !== true ? { subscript: false } : {}) } };
    case 'subscript': return { other: { subscript: context.subscript !== true, ...(context.subscript !== true ? { superscript: false } : {}) } };
    case 'fontSize': return { fontSize: mark.value };
    case 'fontFamily': return { fontFamily: { ascii: mark.value, hAnsi: mark.value } };
    case 'color': return { color: { rgb: mark.value } };
  }
}

export function applyDocxCommand(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  prepared: PreparedDocxCommand,
  context: DocxCommandApplyContext,
  environment?: ResidentCommandHooksEnvironment
): DocxCommandApplyResult {
  const { inputTime, lastInputTime } = prepared.command;
  if (inputTime === undefined) return applyPreparedDocxCommand(engine, state, prepared, context, environment);
  engine.setUndoInputTime([inputTime, lastInputTime ?? inputTime]);
  try {
    return applyPreparedDocxCommand(engine, state, prepared, context, environment);
  } finally {
    engine.setUndoInputTime(null);
  }
}

function applyPreparedDocxCommand(
  engine: ResidentDocxCommandEngine,
  state: ResidentDocxCommandState,
  prepared: PreparedDocxCommand,
  _context: DocxCommandApplyContext,
  environment?: ResidentCommandHooksEnvironment
): DocxCommandApplyResult {
  const { command, range, selection } = prepared;
  const result = (status: 'applied' | 'noop', toolbar?: YrsSelectionContext): DocxCommandApplyResult => {
    if (state.activePart && state.selection && state.selection.focus.story !== state.activePart.story &&
      !state.selection.focus.story.startsWith(`${state.activePart.story}:`)) state.activePart = null;
    const corrected = state.selection ? readCommandSelection(engine, state) : null;
    return { status, selection: corrected?.ok ? corrected.selection : null, editState: projectDocxEditState(engine, state, corrected, toolbar) };
  };
  if (command.kind === 'setRenderEnvironment') {
    const { reservations, ...rest } = command.payload;
    const current = reservations?.version === engine.proposalEngine.version();
    environment!.setRenderEnvironment!(reservations
      ? { ...rest, reservations: current
        ? resolveReservations(engine.geometryReader, reservations, engine.proposalEngine.version(),
          environment?.getLayoutReservations?.() ?? null)
        : retainReservationOwners(reservations, environment?.getLayoutReservations?.() ?? null) }
      : rest, reservations ? (current ? { version: reservations.version, key: reservations.key ?? null } : null) : undefined);
    return result('noop');
  }
  if (command.kind === 'applyEdits') {
    const applied = applyCommandBatch(engine, state, command.payload);
    const projected = result(applied.ok && applied.applied ? 'applied' : 'noop');
    projected.editState!.batch = { seq: command.seq, result: applied };
    return applied.ok ? projected : { ...projected, status: 'refused', reason: applied.failure.code };
  }
  if (command.kind === 'setFontEnvironment') return result('noop');
  if (command.kind === 'setContentControlValue' || command.kind === 'addRepeatingItem' || command.kind === 'removeRepeatingItem') {
    const target = resolveControlHandle(engine, command.payload.handle);
    if (!target) return { ...result('noop'), status: 'refused', reason: 'stale-control' };
    const before = engine.proposalEngine.version();
    engine.beginUndoCapture();
    engine.addUndoBoundary();
    try {
      if (command.kind === 'setContentControlValue') engine.setControlValue(target.at, command.payload.value);
      else {
        engine.repeatControlAt(target.at, command.kind === 'removeRepeatingItem');
        if (state.selection && !engine.geometryReader.hasStory(state.selection.focus.story)) {
          const story = target.at.story.startsWith('body') ? 'body' : target.at.story.split(':').slice(0, 2).join(':');
          const paragraph = engine.geometryReader.paragraphSpans(story)[0];
          if (!paragraph) throw new Error('Control removal has no caret target');
          const at = { story, paraId: paragraph.paraId, offset: 0 };
          engine.setSelection(at, at);
          state.selection = mintCommandSelection(engine, state, { anchor: at, focus: at });
          state.storedMarks = null;
        }
      }
    } catch (error) {
      if (error instanceof WebAssembly.RuntimeError || engine.proposalEngine.version() !== before) throw error;
      return { ...result('noop'), status: 'refused', reason: typeof error === 'string' ? error : error instanceof Error ? error.message : 'invalid-control-value' };
    } finally { engine.addUndoBoundary(); }
    return result(engine.proposalEngine.version() !== before ? 'applied' : 'noop');
  }
  if (command.kind === 'openPart' || command.kind === 'ensurePart' || command.kind === 'removePart') {
    let target = resolvePart(engine, command.payload);
    let changed = false;
    const choose = (story: string) => {
      const paragraph = engine.geometryReader.paragraphSpans(story)[0];
      if (!paragraph) throw new Error('Part has no caret target');
      const at = { story, paraId: paragraph.paraId, offset: 0 };
      engine.setSelection(at, at);
      state.selection = mintCommandSelection(engine, state, { anchor: at, focus: at });
      state.storedMarks = null;
    };
    if (command.kind !== 'openPart') { engine.beginUndoCapture(); engine.addUndoBoundary(); }
    if (command.kind === 'ensurePart' && !target) {
      changed = engine.setPartDocument(ensurePartDocument(engine, command.payload)!);
      target = resolvePart(engine, command.payload);
    }
    if (command.kind === 'removePart') {
      const document = removePartDocument(engine, command.payload);
      if (!document || !target) return { ...result('noop'), status: 'refused', reason: 'unsupported-part' };
      choose('body');
      state.activePart = null;
      changed = engine.setPartDocument(document, [target.story]);
    } else {
      if (!target) throw new Error('Part creation did not resolve');
      choose(target.story);
      state.activePart = target;
      selectResidentPartPage(engine, target.story, target.page);
    }
    if (command.kind !== 'openPart') engine.addUndoBoundary();
    return result(changed ? 'applied' : 'noop');
  }
  if (propertyCommand(command.kind)) return result(applyPropertyCommand(engine, command).status as 'applied' | 'noop');
  if (command.kind === 'setSectionMargins') {
    engine.beginUndoCapture();
    engine.addUndoBoundary();
    const changed = engine.setSectionMargins(command.payload.section, command.payload.margins);
    engine.addUndoBoundary();
    return result(changed ? 'applied' : 'noop');
  }
  if (command.kind === 'resizeObject' || command.kind === 'moveObject') {
    const at = resolveResidentObject(engine, command.payload.handle);
    if (!at) return { ...result('noop'), status: 'refused', reason: 'stale-object' };
    const version = engine.proposalEngine.version();
    if (command.kind === 'resizeObject' && 'widthTwips' in command.payload) {
      if (!at.cell) return { ...result('noop'), status: 'refused', reason: 'stale-object' };
      engine.addUndoBoundary();
      engine.setColumnWidth(at.cell, command.payload.widthTwips);
      engine.addUndoBoundary();
      return result(engine.proposalEngine.version() === version ? 'noop' : 'applied');
    }
    const paragraph = engine.geometryReader.paragraphSpans(at.story).findIndex(span => span.paraId === at.paraId);
    let offset = 0;
    const segment = engine.paragraphSegments(at.story, [paragraph])[0]!.find(part => {
      const start = offset;
      offset += part.kind === 'text' ? part.text.length : 1;
      return start === at.offset;
    });
    if (segment?.kind !== 'embed') return { ...result('noop'), status: 'refused', reason: 'stale-object' };
    const size = command.kind === 'resizeObject' && 'widthEmu' in command.payload ? command.payload : null;
    const widthEmu = size ? size.widthEmu : Number(segment.payload.width ?? 100) * 9525;
    const heightEmu = size ? size.heightEmu : Number(segment.payload.height ?? 100) * 9525;
    const position = segment.payload.position as Record<string, unknown> | null | undefined;
    const move = command.kind === 'moveObject' ? command.payload : size;
    const offsets = move?.hOffsetEmu !== undefined ? {
      horizontal: { relativeTo: 'page', posOffset: move.hOffsetEmu },
      vertical: { relativeTo: 'page', posOffset: move.vOffsetEmu },
    } : null;
    const placement = !offsets ? {}
      : isAnchoredImagePayload(segment.payload) ? { position: { ...position, ...offsets } }
      : { wrapType: 'square', displayMode: 'float', cssFloat: 'none', position: { ...position, behindDoc: false, ...offsets } };
    engine.addUndoBoundary();
    engine.setImageGeometryAt(at, { widthEmu, heightEmu, other: { width: widthEmu / 9525, height: heightEmu / 9525, ...placement } });
    engine.addUndoBoundary();
    return result(engine.proposalEngine.version() === version ? 'noop' : 'applied');
  }
  if (command.kind === 'replaceMatches') {
    if (!command.payload.targets.length) return result('noop');
    engine.beginUndoCapture();
    const author = command.payload.suggesting;
    const applied = engine.proposalEngine.applyEdits({
      expectVersion: command.payload.expectVersion, history: 'separate',
      steps: command.payload.targets.map(target => ({
        op: 'replaceText', target: { kind: 'range', ...target }, text: command.payload.text,
        ...(author ? { suggest: { author: author.name, date: author.date } } : {}),
      })),
    });
    if (!applied.ok) return { ...result('noop'), status: 'refused', reason: applied.failure.code };
    if (applied.applied) state.storedMarks = null;
    return result(applied.applied ? 'applied' : 'noop');
  }
  if (command.kind === 'undo' || command.kind === 'redo') {
    const beforeHead = engine.selection()?.head ?? null;
    const tables = beforeHead ? liveTableAncestors(engine, beforeHead.story) : [];
    const nearby = tables.map((table) => nearTable(engine, table)).filter((at): at is YrsLoc => !!at);
    engine.addUndoBoundary();
    const changed = command.kind === 'undo' ? engine.undo() : engine.redo();
    if (!changed) return result('noop');
    state.storedMarks = null;
    if (state.activePart && !engine.geometryReader.hasStory(state.activePart.story)) {
      const paragraph = engine.geometryReader.paragraphSpans('body')[0];
      if (paragraph) {
        const at = { story: 'body', paraId: paragraph.paraId, offset: 0 };
        engine.setSelection(at, at);
        state.selection = mintCommandSelection(engine, state, { anchor: at, focus: at });
      }
      state.activePart = null;
    }
    const resolved = state.selection ? readCommandSelection(engine, state) : null;
    const story = state.selection?.focus.story;
    const native = !story || engine.geometryReader.hasStory(story) ? engine.selection() : null;
    const selected = resolved?.ok ? resolved.selection : native && commandSelectionRange(engine, { anchor: native.anchor, focus: native.head })
      ? { anchor: native.anchor, focus: native.head } : null;
    const root = story && story !== 'body' && !story.startsWith('body:') ? story.split(':').slice(0, 2).join(':') : 'body';
    const fallbackStory = engine.geometryReader.hasStory(root) ? root : 'body';
    const near = nearby.map((at) => pastLeadingBlocks(engine, at)).find((at) => !!at) ?? null;
    const first = selected ? null : engine.geometryReader.paragraphSpans(fallbackStory)[0];
    const at = first ? { story: fallbackStory, paraId: first.paraId, offset: 0 } : null;
    const repaired = selected ?? (near ? { anchor: near, focus: near } : null) ?? (at ? { anchor: at, focus: at } : null);
    state.selection = repaired ? mintCommandSelection(engine, state, repaired) : null;
    if (state.selection) {
      engine.setSelection(state.selection.anchor, state.selection.focus);
      if (tables.length) syncCellSelectionWithCaret(engine, state.selection.focus.story);
    }
    return result('applied');
  }
  if (Object.hasOwn(reviewCommandKeys, command.kind)) {
    const applied = applyReviewCommand(engine, state, command as ReviewCommand);
    if (applied.status === 'refused') return { ...result('noop'), ...applied };
    repairCommandSelection(engine, state);
    state.storedMarks = null;
    const projected = result(applied.status);
    if (applied.comment) projected.editState!.comment = { seq: command.seq, ...applied.comment };
    if (applied.commentChange) projected.editState!.commentChange = { seq: command.seq, ...applied.commentChange };
    return projected;
  }
  if (command.kind === 'setReadOnly') {
    state.readOnly = command.payload.readOnly;
    return result('noop');
  }
  if (!range || !selection) throw new Error('Prepared command lacks selection');
  const prior = state.selection ? readCommandSelection(engine, state) : null;
  const priorSelection = prior?.ok ? prior.selection : null;
  const publish = (selected: DocxCommandSelection, force = false) => {
    if (force || !sameSelection(state.selection, selected)) {
      engine.setSelection(selected.anchor, selected.focus);
      state.selection = mintCommandSelection(engine, state, selected);
    }
  };
  if (command.kind === 'navigateCell') {
    if (!prepared.cellNavigation) throw new Error('Prepared command lacks cell navigation');
    const selected = applyCellNavigation(engine, prepared.cellNavigation, command.payload);
    state.storedMarks = null;
    publish(selected, true);
    return result(prepared.cellNavigation.insertRow ? 'applied' : 'noop');
  }
  if (command.kind === 'select' || command.kind === 'navigate' || command.kind === 'selectPoint' ||
    command.kind === 'extendSelection' || command.kind === 'selectTableCells') {
    if (!sameSelection(priorSelection, selection)) state.storedMarks = null;
    if (prepared.cells !== undefined) writeCellSelection(engine, prepared.cells);
    publish(selection, true);
    return result('noop');
  }
  if (state.storedMarks && (state.storedMarks.story !== selection.focus.story ||
    state.storedMarks.paraId !== selection.focus.paraId || !sameSelection(priorSelection, selection) ||
    selection.anchor.paraId !== selection.focus.paraId || selection.anchor.offset !== selection.focus.offset)) state.storedMarks = null;
  publish(selection);
  if (collapsed(range) && ['toggleMark', 'toggleStrike', 'formatRange', 'formatSelection', 'clearFormatting'].includes(command.kind)) {
    const current = state.storedMarks;
    const toolbar = command.kind === 'toggleMark' || command.kind === 'toggleStrike' ? engine.selectionContext(range) : undefined;
    const delta = command.kind === 'toggleStrike' ? { strike: withStoredMarks(toolbar!, current).strike !== true } : command.kind === 'toggleMark' ? markDelta(command.payload.mark, withStoredMarks(toolbar!, current)) :
      command.kind === 'formatRange' || command.kind === 'formatSelection' ? command.payload.delta : {};
    state.storedMarks = {
      story: range.story, paraId: range.start.paraId,
      clear: command.kind === 'clearFormatting' || (current?.clear ?? false),
      delta: command.kind === 'clearFormatting' ? {} : {
        ...current?.delta, ...delta,
        ...(current?.delta.other || delta.other ? { other: { ...current?.delta.other, ...delta.other } } : {}),
      },
    };
    return result('noop', toolbar);
  }
  resetCellSelectionForEdit(engine, state, command.kind);
  if ((command.kind === 'insertText' || command.kind === 'replaceSelection') && command.payload.text === '' && collapsed(range)) return result('noop');
  if ((command.kind === 'deleteBackward' || command.kind === 'deleteForward') && collapsed(range)) return result('noop');
  if ((command.kind === 'formatRange' || command.kind === 'formatSelection') && Object.keys(command.payload.delta).length === 0) return result('noop');
  if (command.kind === 'setParagraphAttrs' && Object.keys(command.payload.attrs).length === 0) return result('noop');
  const version = engine.proposalEngine.version();
  engine.beginUndoCapture();
  const discrete = ['setTabStops', 'removeTabStop', 'toggleList', 'changeIndent', 'insertImage', 'insertTable', 'pasteFragment', 'insertPageBreak', 'insertSectionBreak', 'lineBreak'].includes(command.kind) ||
    isTableCommandKind(command.kind);
  if (discrete) engine.addUndoBoundary();
  const caret = (at: YrsLoc) => publish({ anchor: at, focus: at }, true);
  switch (command.kind) {
    case 'setTabStops':
      engine.setParagraphAttrs(range, { tabs: command.payload.tabs });
      break;
    case 'removeTabStop': {
      const paragraphs = engine.paragraphs(range.story);
      const start = paragraphs.findIndex(paragraph => paragraph.paraId === range.start.paraId);
      const end = paragraphs.findIndex(paragraph => paragraph.paraId === range.end.paraId);
      for (const paragraph of paragraphs.slice(start, end + 1)) {
        const tabs = nativeTabStops(paragraph.properties.tabs);
        if (!tabs.some(tab => tab.position === command.payload.positionTwips)) continue;
        const at = { paraId: paragraph.paraId, offset: 0 };
        engine.setParagraphAttrs({ story: range.story, start: at, end: at }, { tabs: tabs.filter(tab => tab.position !== command.payload.positionTwips) });
      }
      break;
    }
    case 'insertText': case 'replaceSelection': {
      const receipt = collapsed(range)
        ? engine.insertText({ story: range.story, ...range.start }, command.payload.text, command.payload.suggesting)
        : engine.replaceRange(range, command.payload.text, command.payload.suggesting);
      if (!receipt.range) throw new Error('Text receipt lacks range');
      const stored = state.storedMarks;
      if (stored && command.payload.text && !collapsed(receipt.range)) {
        if (stored.clear) engine.clearFormatting(receipt.range);
        if (Object.keys(stored.delta).length) engine.formatRange(receipt.range, stored.delta);
      }
      caret({ story: receipt.range.story, ...receipt.range.end });
      break;
    }
    case 'deleteBackward': case 'deleteForward': {
      const receipt = engine.deleteRange(range, command.payload.suggesting);
      if (!receipt.range) throw new Error('Delete receipt lacks range');
      if (state.storedMarks && (state.storedMarks.story !== receipt.range.story ||
        state.storedMarks.paraId !== receipt.range.start.paraId)) state.storedMarks = null;
      const advance = command.kind === 'deleteForward' && command.payload.suggesting !== undefined &&
        selection.anchor.story === selection.focus.story && selection.anchor.paraId === selection.focus.paraId &&
        selection.anchor.offset === selection.focus.offset;
      caret({ story: receipt.range.story, ...(advance ? receipt.range.end : receipt.range.start) });
      break;
    }
    case 'lineBreak': {
      let at = { story: range.story, ...range.start };
      if (!collapsed(range)) {
        const removed = engine.deleteRange(range, command.payload.suggesting);
        if (!removed.range) throw new Error('Delete receipt lacks range');
        at = { story: removed.range.story, ...removed.range.start };
      }
      const receipt = engine.insertLineBreak(at, command.payload.suggesting);
      if (!receipt.range) throw new Error('Line break receipt lacks range');
      const stored = state.storedMarks;
      if (stored) {
        if (stored.clear) engine.clearFormatting(receipt.range);
        if (Object.keys(stored.delta).length) engine.formatRange(receipt.range, stored.delta);
      }
      caret({ story: receipt.range.story, ...receipt.range.end });
      break;
    }
    case 'paragraphBreak': {
      let at = { story: range.story, ...range.start };
      if (!collapsed(range)) {
        const receipt = engine.deleteRange(range, command.payload.suggesting);
        if (!receipt.range) throw new Error('Delete receipt lacks range');
        at = { story: receipt.range.story, ...receipt.range.start };
      }
      const receipt = engine.splitParagraph(at, command.payload.suggesting);
      if (state.storedMarks) state.storedMarks = { ...state.storedMarks, story: at.story, paraId: receipt.secondParaId };
      caret({ story: at.story, paraId: receipt.secondParaId, offset: 0 });
      break;
    }
    case 'toggleList': case 'changeIndent':
      for (const change of prepared.paragraphChanges ?? []) engine.setParagraphAttrs(change.range, change.attrs, command.payload.suggesting);
      break;
    case 'insertImage': case 'insertTable': case 'insertPageBreak': case 'insertSectionBreak': {
      let at = { story: range.story, ...range.start };
      if (!collapsed(range)) {
        const removed = engine.deleteRange(range, command.payload.suggesting);
        if (!removed.range) throw new Error('Delete receipt lacks range');
        at = { story: removed.range.story, ...removed.range.start };
      }
      state.storedMarks = null;
      if (command.kind === 'insertImage') {
        const receipt = engine.insertImage(at, command.payload.image, command.payload.suggesting);
        if (!receipt.range) throw new Error('Image receipt lacks range');
        caret({ story: receipt.range.story, ...receipt.range.end });
      } else if (command.kind === 'insertTable') {
        const receipt = engine.insertTable(at, command.payload.rows, command.payload.columns, command.payload.suggesting);
        const story = receipt.createdStoryIds[0];
        const paragraph = story ? engine.paragraphs(story)[0] : undefined;
        if (!story || !paragraph) throw new Error('Table receipt lacks cell');
        const cell = { ...receipt.table, row: 0, column: 0 };
        engine.setCellSelection({ anchor: cell, head: cell });
        caret({ story, paraId: paragraph.paraId, offset: 0 });
      } else {
        if (command.kind === 'insertPageBreak') {
          const target = engine.insertPageBreak(at);
          caret({ ...target, offset: target.offset + 1 });
        } else {
          const properties = collapsed(range) ? prepared.sectionBreakProperties : residentSectionPropertiesAt(engine, at);
          if (!properties) throw new Error('Prepared command lacks section properties');
          caret(engine.splitSection(at, properties, command.payload.type));
        }
      }
      break;
    }
    case 'insertTableRow': case 'insertTableColumn': case 'deleteTableRow': case 'deleteTableColumn':
    case 'deleteTable': case 'mergeTableCells': case 'splitTableCell': case 'setTableCellShading': case 'setTableBorders': case 'setTableProperties': {
      if (!prepared.table) throw new Error('Prepared command lacks table target');
      const focus = selection.focus;
      applyTableCommand(engine, state, prepared.command as TableCommandIntent, prepared.table,
        { story: focus.story, paraId: focus.paraId, offset: focus.offset },
        { caret, repair: () => repairCommandSelection(engine, state) });
      break;
    }
    case 'toggleStrike': engine.formatRange(range, { strike: engine.selectionContext(range).strike !== true }); break;
    case 'pasteFragment': {
      state.storedMarks = null;
      caret(applyClipboardFragment(engine, range, command.payload.fragment, command.payload.suggesting));
      break;
    }
    case 'toggleMark': engine.toggleMark(range, command.payload.mark); break;
    case 'formatSelection': case 'formatRange': engine.formatRange(range, command.payload.delta); break;
    case 'clearFormatting': engine.clearFormatting(range); break;
    case 'setParagraphAttrs': engine.setParagraphAttrs(range, paragraphAttrs(command.payload.attrs), command.payload.suggesting); break;
    case 'applyParagraphStyle': engine.applyParagraphStyle(range, command.payload.styleId, command.payload.suggesting); break;
  }
  resetCellSelectionForEdit(engine, state, command.kind);
  if (discrete) engine.addUndoBoundary();
  return result(engine.proposalEngine.version() === version ? 'noop' : 'applied');
}

export const createResidentDocxCommandHooks: ResidentCommandHooksFactory = (engine, init, environment) => {
  const recovery = init.recoveryState;
  const native = engine.selection();
  const first = !native && init.selection === undefined && engine.geometryReader.hasStory('body')
    ? engine.geometryReader.paragraphSpans('body')[0] : undefined;
  const fallback = first ? { story: 'body', paraId: first.paraId, offset: 0 } : null;
  const state: ResidentDocxCommandState = {
    readOnly: init.readOnly,
    selection: recovery ? structuredClone(recovery.selection) : init.selection !== undefined ? init.selection : native
      ? { anchor: native.anchor, focus: native.head } : fallback ? { anchor: fallback, focus: fallback } : null,
    storedMarks: structuredClone(recovery?.storedMarks ?? null),
    activePart: structuredClone(recovery?.activePart),
    stickyPositions: new Map(structuredClone(recovery?.stickyPositions ?? [])),
    proposals: environment?.proposals ?? createProposalRegistry(engine.proposalEngine),
  };
  if (recovery) {
    state.proposals!.mirror({ version: engine.proposalEngine.version(), proposals: recovery.proposals });
    state.proposals!.mirror(null);
  }
  if (state.selection) {
    const selected = {
      anchor: { story: state.selection.anchor.story, paraId: state.selection.anchor.paraId, offset: state.selection.anchor.offset },
      focus: { story: state.selection.focus.story, paraId: state.selection.focus.paraId, offset: state.selection.focus.offset },
    };
    if (!commandSelectionRange(engine, selected)) throw new Error('Invalid initial command selection');
    engine.setSelection(selected.anchor, selected.focus);
    state.selection = mintCommandSelection(engine, state, selected);
  }
  if (recovery?.cellSelection) engine.setCellSelection(recovery.cellSelection);
  registerCommandAnchorState(engine, state);
  let projection: { version: string; selection: DocxCommandSelection | null; editState: ReturnType<typeof projectDocxEditState> } | null = null;
  const remember = (result: DocxCommandApplyResult) => {
    projection = { version: `${engine.proposalEngine.version()}:${state.proposals?.snapshot().previewVersion}`, selection: result.selection ?? null, editState: result.editState! };
    return result;
  };
  const projected = () => {
    if (state.activePart && !engine.geometryReader.hasStory(state.activePart.story)) {
      const paragraph = engine.geometryReader.paragraphSpans('body')[0];
      if (paragraph) {
        const at = { story: 'body', paraId: paragraph.paraId, offset: 0 };
        engine.setSelection(at, at);
        state.selection = mintCommandSelection(engine, state, { anchor: at, focus: at });
      }
      state.activePart = null;
    }
    if (!projection || projection.version !== `${engine.proposalEngine.version()}:${state.proposals?.snapshot().previewVersion}`) {
      const resolved = state.selection ? readCommandSelection(engine, state) : null;
      remember({ status: 'noop', selection: resolved?.ok ? resolved.selection : null,
        editState: projectDocxEditState(engine, state, resolved) });
    }
    return projection!;
  };
  return {
    preflight: (command, context) => {
      if (command.basis.epoch !== init.epoch && commandPolicy(command.kind)?.origin !== 'remote') {
        return { status: 'refused', reason: 'stale-epoch', selection: projected().selection, editState: projected().editState };
      }
      const result = preflightDocxCommand(engine, state, command, context, environment);
      if (result.status === 'refused') remember(result);
      return result;
    },
    apply: (command, context) => remember(applyDocxCommand(engine, state, command, context, environment)),
    afterRemoteMerge: (context) => {
      state.proposals?.reconcile();
      const run = state.remoteMerges;
      const storedMarks = state.storedMarks;
      const resolved = state.selection ? readCommandSelection(engine, state) : null;
      repairCommandSelection(engine, state);
      const survived = resolved?.ok && sameSelection(state.selection, resolved.selection);
      state.remoteMerges = survived
        ? { from: run?.to === context.currentVersion ? run.from : context.currentVersion, to: context.currentVersion + 1 }
        : undefined;
      state.storedMarks = survived && storedMarks && state.selection && sameSelection({ anchor: state.selection.anchor, focus: state.selection.anchor }, state.selection)
        ? { ...storedMarks, story: state.selection.focus.story, paraId: state.selection.focus.paraId }
        : null;
    },
    selection: () => projected().selection,
    editState: () => projected().editState,
    recoveryState: () => ({
      selection: projected().selection, storedMarks: state.storedMarks, activePart: state.activePart,
      stickyPositions: [...state.stickyPositions], cellSelection: readCellSelection(engine),
      proposals: state.proposals!.exportState(),
    }),
  };
};
