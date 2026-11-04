import type { TranslationKey } from '@betteroffice/pptx-i18n';
import {
  commandReason,
  documentGate,
  reviewGate,
  slideGate,
  textGate,
  writeGate,
  type PptxCommandEnvironment,
  type PptxCommandEvaluation,
  type PptxMenuReasonCode,
} from './evaluate';
import type { CommandReason, PptxCommandId } from './types';

const REASON_KEYS: Record<PptxMenuReasonCode, TranslationKey> = {
  'clipboard-unavailable': 'commands.reasons.clipboardUnavailable',
  'table-required': 'commands.reasons.tableRequired',
  'last-slide': 'commands.reasons.lastSlide',
};

const MENU_COMMANDS = new Set<PptxCommandId>(['cut', 'copy', 'paste', 'deleteObject', 'deleteSlide',
  'duplicateSlide', 'tableInsert', 'tableDelete']);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function validArgs(id: PptxCommandId, args: unknown): boolean {
  if (args === undefined) return true;
  if (id === 'tableInsert') {
    return isRecord(args) && (args.kind === 'row' || args.kind === 'column') &&
      (args.position === 'before' || args.position === 'after');
  }
  if (id === 'tableDelete') return isRecord(args) && ['row', 'column', 'table'].includes(args.kind as string);
  return args === null;
}

function evaluation(gate: PptxCommandEvaluation['gate'], reasonKey?: TranslationKey): PptxCommandEvaluation {
  const key = gate && gate in REASON_KEYS ? REASON_KEYS[gate as PptxMenuReasonCode] : reasonKey;
  return key ? { gate, reasonKey: key } : { gate };
}

/** @internal Why a menu command cannot run, with its message from the menu table. */
export function menuReason(code: PptxMenuReasonCode,
  env: Pick<PptxCommandEnvironment, 'translate'>): CommandReason<PptxMenuReasonCode> {
  return commandReason(code, env, REASON_KEYS[code]);
}

/** @internal The result for a clipboard call the browser refused, or null for any other error. */
export function clipboardRefusal(error: unknown,
  env: Pick<PptxCommandEnvironment, 'translate'>): { ok: false; failure: CommandReason<PptxMenuReasonCode> } | null {
  return error && typeof error === 'object' && 'name' in error && error.name === 'NotAllowedError'
    ? { ok: false, failure: menuReason('clipboard-unavailable', env) } : null;
}

/**
 * @internal Gates for the clipboard, object, slide and table commands the
 * context menus run; other ids are left to the built-in gate.
 */
export function pptxMenuGates(id: PptxCommandId, args: unknown,
  env: PptxCommandEnvironment): PptxCommandEvaluation | undefined {
  if (!MENU_COMMANDS.has(id)) return undefined;
  if (!validArgs(id, args)) return { gate: documentGate(env) ?? 'invalid-arguments' };
  const current = env.text && env.text !== 'unsupported' ? env.text : null;
  switch (id) {
    case 'cut':
    case 'copy':
      return evaluation((id === 'cut' ? writeGate(env) : documentGate(env)) ?? reviewGate(env) ??
        (env.text === 'unsupported' ? 'unsupported-selection' : null) ??
        (current?.kind === 'range' ? null : 'text-selection-required') ??
        (env.clipboardWrite ? null : 'clipboard-unavailable'));
    case 'paste':
      return evaluation(textGate(env) ?? (current?.kind === 'shape' ? 'text-selection-required' : null) ??
        (env.clipboardRead ? null : 'clipboard-unavailable'));
    case 'deleteObject':
      return evaluation(writeGate(env) ?? reviewGate(env) ?? (env.shape ? null : 'shape-required'),
        env.shape ? undefined : 'commands.reasons.objectRequired');
    case 'deleteSlide':
    case 'duplicateSlide':
      return evaluation(writeGate(env) ?? reviewGate(env) ?? slideGate(env) ??
        (id === 'deleteSlide' && (env.slideCount ?? 0) <= 1 ? 'last-slide' : null));
    default:
      return evaluation(writeGate(env) ?? reviewGate(env) ?? slideGate(env) ??
        (env.table ? null : 'table-required'));
  }
}
