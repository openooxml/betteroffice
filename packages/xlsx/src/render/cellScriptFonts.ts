import { resolveScriptFallbackFace, type BundledFontFace, type BundledFontScript } from '@betteroffice/fonts';
import { textScripts } from '../../../../shared/office-session/textScripts';
import type { DisplayList, TextCmd } from '../display-list/types';
import { runPieces } from './textRuns';

interface ScriptFontRequest {
  fontFamily: string;
  bold: boolean;
  italic: boolean;
  script: BundledFontScript;
  face: BundledFontFace;
}

function cellScripts(cmd: TextCmd & { cssFontFamily?: string }): BundledFontScript[] {
  if (cmd.chart || cmd.cssFontFamily !== undefined) return [];
  const text = (cmd.text ?? '') + (cmd.numberFormat?.parts.map(({ character }) => character).join('') ?? '');
  return /^[\u0000-\u02ff]*$/u.test(text) ? [] : textScripts(text) as BundledFontScript[];
}

export function cellScriptFaces(cmd: TextCmd, scripts = cellScripts(cmd)): ScriptFontRequest[] {
  const { bold = false, italic = false } = cmd;
  const fontFamily = cmd.fontFamily ?? '';
  return scripts.flatMap((script) => {
    const face = resolveScriptFallbackFace(script, bold, italic, fontFamily);
    return face ? [{ fontFamily, bold, italic, script, face }] : [];
  });
}

export function cellScriptFontRequests(displayList: DisplayList): ScriptFontRequest[] {
  const requests = new Map<string, ScriptFontRequest>();
  for (const cmd of displayList.commands) {
    if (cmd.op !== 'text') continue;
    const scripts = cellScripts(cmd);
    if (!scripts.length) continue;
    const pieces = runPieces(cmd);
    const styles = pieces ? pieces.map((piece) => piece.cmd) : [cmd];
    for (const style of styles) {
      for (const request of cellScriptFaces(style, scripts)) requests.set(request.face.file, request);
    }
  }
  return [...requests.values()];
}
