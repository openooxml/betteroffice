import type { BundledFontScript } from '@betteroffice/fonts';
import type { DisplayList, TextCmd } from '../../packages/xlsx/src/display-list/types';
import { nonCjkScriptUnicodeRange, textScripts } from './textScripts';
import { runPieces, textWithNumberFormatParts } from './xlsxTextRuns';

export interface CellScriptFontRequest {
  fontFamily: string;
  bold: boolean;
  italic: boolean;
  script: BundledFontScript;
}

export type CellScriptFamilies = (cmd: TextCmd) => string[] | undefined;
export type ScriptFontDisplayList = DisplayList & { cellScriptFamilies?: CellScriptFamilies };
export type CellScriptCache = Map<string, readonly BundledFontScript[]>;

const NO_SCRIPTS: readonly BundledFontScript[] = [];
const SCRIPT_CACHE_CAPACITY = 1024;
const SCRIPT_CACHE_MAX_TEXT_LENGTH = 512;
const scriptsByText = new Map<string, readonly BundledFontScript[]>();

function scriptsForText(text: string): readonly BundledFontScript[] {
  if (/^[\u0000-\u02ff]*$/u.test(text)) return NO_SCRIPTS;
  const known = scriptsByText.get(text);
  if (known) return known;
  const scripts = textScripts(text).filter((script) => nonCjkScriptUnicodeRange(script)) as BundledFontScript[];
  if (text.length <= SCRIPT_CACHE_MAX_TEXT_LENGTH) {
    if (scriptsByText.size >= SCRIPT_CACHE_CAPACITY) scriptsByText.delete(scriptsByText.keys().next().value!);
    scriptsByText.set(text, scripts);
  }
  return scripts;
}

function cellScripts(cmd: TextCmd, frameScripts?: CellScriptCache): readonly BundledFontScript[] {
  const text = textWithNumberFormatParts(cmd);
  const known = frameScripts?.get(text);
  if (known) return known;
  const scripts = scriptsForText(text);
  frameScripts?.set(text, scripts);
  return scripts;
}

export function cellFontKey(fontFamily: string, bold: boolean, italic: boolean): string {
  return `${Number(bold)}${Number(italic)}\0${fontFamily.trim().toLowerCase()}`;
}

export function cellScriptFontKey({ script, fontFamily, bold, italic }: CellScriptFontRequest): string {
  return `${script}\0${cellFontKey(fontFamily, bold, italic)}`;
}

export function cellScriptFamilies(cmd: TextCmd, families: ReadonlyMap<string, string>, frameScripts?: CellScriptCache): string[] | undefined {
  if (cmd.chart) return undefined;
  const scripts = cellScripts(cmd, frameScripts);
  if (!scripts.length) return undefined;
  const style = cellFontKey(cmd.fontFamily ?? '', cmd.bold === true, cmd.italic === true);
  const names: string[] = [];
  for (const script of scripts) {
    const name = families.get(`${script}\0${style}`);
    if (name !== undefined) names.push(name);
  }
  return names;
}

export function cellScriptFontRequests(displayList: DisplayList, frameScripts?: CellScriptCache): CellScriptFontRequest[] {
  const requests = new Map<string, CellScriptFontRequest>();
  const add = (cmd: TextCmd, scripts: readonly BundledFontScript[]) => {
    if (!scripts.length) return;
    const fontFamily = cmd.fontFamily ?? '';
    const bold = cmd.bold === true;
    const italic = cmd.italic === true;
    const style = cellFontKey(fontFamily, bold, italic);
    for (const script of scripts) {
      const key = `${script}\0${style}`;
      if (!requests.has(key)) requests.set(key, { script, fontFamily, bold, italic });
    }
  };
  for (const cmd of displayList.commands) {
    if (cmd.op !== 'text' || cmd.chart) continue;
    const pieces = runPieces(cmd);
    if (pieces) {
      for (const piece of pieces) add(piece.cmd, cellScripts(piece.cmd, frameScripts));
    } else add(cmd, cellScripts(cmd, frameScripts));
  }
  return [...requests.values()];
}
