import type { BundledFontFace, BundledFontScript, resolveScriptFallbackFace } from '@betteroffice/fonts';
import type { DisplayList, TextCmd } from '../../packages/xlsx/src/display-list/types';
import { nonCjkScriptUnicodeRange, textScripts } from './textScripts';
import { runPieces } from './xlsxTextRuns';

export interface CellScriptFontRequest {
  fontFamily: string;
  bold: boolean;
  italic: boolean;
  script: BundledFontScript;
}

const SCRIPT_FONTS = Symbol.for('betteroffice.xlsx.scriptFonts');
type Resolver = typeof resolveScriptFallbackFace;
const scope = globalThis as typeof globalThis & {
  [SCRIPT_FONTS]?: {
    resolve?: Resolver;
    faces: Map<string, BundledFontFace | undefined>;
    scripts: WeakMap<TextCmd, { text: string; scripts: readonly BundledFontScript[] }>;
  };
};

function cache() {
  return scope[SCRIPT_FONTS] ??= { faces: new Map(), scripts: new WeakMap() };
}

function cellScripts(cmd: TextCmd & { cssFontFamily?: string }): readonly BundledFontScript[] {
  if (cmd.chart || cmd.cssFontFamily !== undefined) return [];
  const { scripts } = cache();
  const text = (cmd.text ?? '') + (cmd.numberFormat?.parts.map(({ character }) => character).join('') ?? '');
  const known = scripts.get(cmd);
  if (known?.text === text) return known.scripts;
  const detected = /^[\u0000-\u02ff]*$/u.test(text) ? []
    : textScripts(text).filter((script) => nonCjkScriptUnicodeRange(script)) as BundledFontScript[];
  scripts.set(cmd, { text, scripts: detected });
  return detected;
}

function requestFor(script: BundledFontScript, { fontFamily, bold = false, italic = false }: TextCmd): CellScriptFontRequest {
  return { script, fontFamily: fontFamily ?? '', bold, italic };
}

export function cellScriptFontKey({ script, fontFamily, bold, italic }: CellScriptFontRequest): string {
  return JSON.stringify([script, fontFamily.trim().toLowerCase(), bold, italic]);
}

export function cellScriptFace(request: CellScriptFontRequest, resolver?: Resolver): BundledFontFace | undefined {
  const state = cache();
  if (resolver && resolver !== state.resolve) {
    state.resolve = resolver;
    state.faces.clear();
  }
  if (!state.resolve) return undefined;
  const key = cellScriptFontKey(request);
  if (!state.faces.has(key)) {
    const { script, fontFamily, bold, italic } = request;
    state.faces.set(key, state.resolve(script, bold, italic, fontFamily));
  }
  return state.faces.get(key);
}

export function cellScriptFaces(cmd: TextCmd): BundledFontFace[] {
  return cellScripts(cmd).flatMap((script) => {
    const face = cellScriptFace(requestFor(script, cmd));
    return face ? [face] : [];
  });
}

export function cellScriptFontRequests(displayList: DisplayList): CellScriptFontRequest[] {
  const requests = new Map<string, CellScriptFontRequest>();
  for (const cmd of displayList.commands) {
    if (cmd.op !== 'text') continue;
    const scripts = cellScripts(cmd);
    if (!scripts.length) continue;
    const pieces = runPieces(cmd);
    const styles = pieces ? pieces.map((piece) => piece.cmd) : [cmd];
    for (const style of styles) {
      for (const script of scripts) {
        const request = requestFor(script, style);
        requests.set(cellScriptFontKey(request), request);
      }
    }
  }
  return [...requests.values()];
}
