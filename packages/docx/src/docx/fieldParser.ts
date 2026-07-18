/** Public field model helpers and a small XmlElement compatibility adapter. */

import type {
  Field,
  FieldType,
  Hyperlink,
  Run,
  SimpleField,
  Theme,
} from '../types/document';
import type { StyleMap } from './styleParser';
import { findChildren, getAttribute, getTextContent, type XmlElement } from './xmlParser';

// Keep the public list explicit; field instructions remain inert data.
export const KNOWN_FIELD_TYPES: FieldType[] = [
  'PAGE', 'NUMPAGES', 'NUMWORDS', 'NUMCHARS', 'DATE', 'TIME', 'CREATEDATE', 'SAVEDATE',
  'PRINTDATE', 'EDITTIME', 'AUTHOR', 'TITLE', 'SUBJECT', 'KEYWORDS', 'COMMENTS', 'FILENAME',
  'FILESIZE', 'TEMPLATE', 'REVNUM', 'DOCPROPERTY', 'DOCVARIABLE', 'REF', 'PAGEREF', 'NOTEREF',
  'HYPERLINK', 'TOC', 'TOA', 'INDEX', 'SEQ', 'STYLEREF', 'AUTONUM', 'AUTONUMLGL',
  'AUTONUMOUT', 'SECTION', 'SECTIONPAGES', 'USERADDRESS', 'USERNAME', 'USERINITIALS',
  'FORMTEXT', 'FORMCHECKBOX', 'FORMDROPDOWN', 'CITATION', 'BIBLIOGRAPHY', 'IF', 'MERGEFIELD',
  'NEXT', 'NEXTIF', 'ASK', 'SET', 'QUOTE', 'INCLUDETEXT', 'INCLUDEPICTURE', 'SYMBOL', 'ADVANCE',
];

export function parseFieldType(instruction: string): FieldType {
  const match = instruction.trim().match(/^\\?([A-Z][A-Z0-9]*)/i);
  const name = match?.[1]?.toUpperCase();
  return name && KNOWN_FIELD_TYPES.includes(name as FieldType) ? (name as FieldType) : 'UNKNOWN';
}

export function isKnownFieldType(type: string): type is FieldType {
  return KNOWN_FIELD_TYPES.includes(type as FieldType);
}

export interface ParsedFieldInstruction {
  type: FieldType;
  raw: string;
  argument?: string;
  switches: FieldSwitch[];
}

export interface FieldSwitch {
  switch: string;
  value?: string;
}

export function parseFieldInstruction(instruction: string): ParsedFieldInstruction {
  const trimmed = instruction.trim();
  const name = trimmed.match(/^\\?([A-Z][A-Z0-9]*)/i)?.[0] ?? '';
  const remaining = trimmed.slice(name.length).trim();
  const switches: FieldSwitch[] = [];
  const positions: number[] = [];
  const pattern = /\\(\*|@|#|!|[a-z])\s*(?:"([^"]*)"|([^\s]*))?/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(remaining))) {
    switches.push({ switch: match[1], ...((match[2] || match[3]) ? { value: match[2] || match[3] } : {}) });
    positions.push(match.index);
  }
  const rawArgument = remaining.slice(0, positions[0] ?? remaining.length).trim();
  const argument =
    rawArgument.startsWith('"') && rawArgument.endsWith('"')
      ? rawArgument.slice(1, -1)
      : rawArgument || undefined;
  return { type: parseFieldType(instruction), raw: instruction, ...(argument ? { argument } : {}), switches };
}

export function getFormatSwitch(instruction: ParsedFieldInstruction): string | undefined {
  return instruction.switches.find((entry) => entry.switch === '*' || entry.switch === '@')?.value;
}

export function hasMergeFormat(instruction: ParsedFieldInstruction): boolean {
  return instruction.switches.find((entry) => entry.switch === '*')?.value?.toUpperCase() === 'MERGEFORMAT';
}

export function parseSimpleField(
  node: XmlElement,
  _styles: StyleMap | null,
  _theme: Theme | null
): SimpleField {
  const instruction = getAttribute(node, 'w', 'instr') ?? '';
  const content: Run[] = findChildren(node, 'w', 'r').map((run) => {
    const text = getTextContent(run);
    return { type: 'run', content: text ? [{ type: 'text', text }] : [] };
  });
  const field: SimpleField = {
    type: 'simpleField',
    instruction,
    fieldType: parseFieldType(instruction),
    content,
  };
  if (/^(1|true)$/i.test(getAttribute(node, 'w', 'fldLock') ?? '')) field.fldLock = true;
  if (/^(1|true)$/i.test(getAttribute(node, 'w', 'dirty') ?? '')) field.dirty = true;
  if (content.length) {
    field.structuredResult = { inline: content };
    field.fieldTree = { version: 1, result: field.structuredResult, displayMode: 'result' };
  }
  return field;
}

export type ComplexFieldState = 'outside' | 'code' | 'result';

export interface ComplexFieldContext {
  state: ComplexFieldState;
  instruction: string;
  codeRuns: Run[];
  resultRuns: Run[];
  fldLock: boolean;
  dirty: boolean;
  nestingLevel: number;
}

export function createComplexFieldContext(): ComplexFieldContext {
  return {
    state: 'outside',
    instruction: '',
    codeRuns: [],
    resultRuns: [],
    fldLock: false,
    dirty: false,
    nestingLevel: 0,
  };
}

export function getFieldDisplayValue(field: Field): string {
  const runs = field.type === 'simpleField' ? field.content : field.fieldResult;
  return runs.map(inlineText).join('');
}

function inlineText(item: Run | Hyperlink): string {
  if (item.type === 'run') return runText(item);
  return item.children.filter((child): child is Run => child.type === 'run').map(runText).join('');
}

function runText(run: Run): string {
  return run.content
    .filter((content) => content.type === 'text')
    .map((content) => content.text)
    .join('');
}

export function isPageNumberField(field: Field): boolean {
  return field.fieldType === 'PAGE';
}

export function isTotalPagesField(field: Field): boolean {
  return field.fieldType === 'NUMPAGES';
}

export function isDateTimeField(field: Field): boolean {
  return ['DATE', 'TIME', 'CREATEDATE', 'SAVEDATE', 'PRINTDATE', 'EDITTIME'].includes(field.fieldType);
}

export function isDocPropertyField(field: Field): boolean {
  return [
    'AUTHOR', 'TITLE', 'SUBJECT', 'KEYWORDS', 'COMMENTS', 'FILENAME', 'FILESIZE', 'TEMPLATE',
    'REVNUM', 'DOCPROPERTY', 'DOCVARIABLE',
  ].includes(field.fieldType);
}

export function isReferenceField(field: Field): boolean {
  return ['REF', 'PAGEREF', 'NOTEREF'].includes(field.fieldType);
}

export function isMergeField(field: Field): boolean {
  return ['MERGEFIELD', 'IF', 'NEXT', 'NEXTIF', 'ASK', 'SET'].includes(field.fieldType);
}

export function isTocField(field: Field): boolean {
  return ['TOC', 'TOA', 'INDEX'].includes(field.fieldType);
}
