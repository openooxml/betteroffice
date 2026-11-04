import type { CellAddr } from '@betteroffice/xlsx';

const MAX_ROWS = 1_048_576;
const MAX_COLS = 16_384;
const MAX_FORMULA_BYTES = 32_768;
const SPACE = '\t\n\v\f\r \u0085   -     　';
const IS_SPACE = new RegExp(`^[${SPACE}]$`);
const LEADING = new RegExp(`^[${SPACE}]+`);
const TRAILING = new RegExp(`[${SPACE}]+$`);
const SEPARATORS = new Set(['+', '-', '*', '/', '^', '&', '%', '=', '<', '>', '(', ')', ',', ';', '{', '}', '@', '#']);

type Ref = { row: number; col: number; absRow: boolean; absCol: boolean };

const isSpace = (character: string) => IS_SPACE.test(character);
const trimStart = (text: string) => text.replace(LEADING, '');
const trimEnd = (text: string) => text.replace(TRAILING, '');
const trim = (text: string) => trimEnd(trimStart(text));
const asciiUpper = (text: string) => text.replace(/[a-z]+/g, (letters) => letters.toUpperCase());
const separator = (character: string) => isSpace(character) || SEPARATORS.has(character);

function utf8Length(text: string): number {
  let bytes = 0;
  for (let index = 0; index < text.length; index++) {
    const code = text.charCodeAt(index);
    if (code < 0x80) bytes += 1;
    else if (code < 0x800) bytes += 2;
    else if (code >= 0xd800 && code < 0xdc00 && (text.charCodeAt(index + 1) & 0xfc00) === 0xdc00) {
      bytes += 4;
      index++;
    } else bytes += 3;
  }
  return bytes;
}

class Output {
  text = '';
  private bytes = 0;

  push(part: string): void {
    const bytes = utf8Length(part);
    if (bytes > MAX_FORMULA_BYTES - this.bytes) throw tooLong();
    this.text += part;
    this.bytes += bytes;
  }
}

const tooLong = () => new Error('formula exceeds the length limit');

function skipQuoted(source: string, start: number, quote: string): number {
  let index = start + 1;
  while (index < source.length) {
    if (source[index] === quote) {
      index += 1;
      if (source[index] !== quote) return index;
    }
    index += 1;
  }
  return source.length;
}

function skipBrackets(source: string, start: number): number {
  let depth = 1;
  let index = start + 1;
  while (index < source.length) {
    const character = source[index];
    if (character === "'") index += index + 1 < source.length ? 1 : 0;
    else if (character === '[') depth += 1;
    else if (character === ']') {
      depth -= 1;
      if (depth === 0) return index + 1;
    }
    index += 1;
  }
  return source.length;
}

function operandEnd(source: string, start: number): number {
  let index = start;
  while (index < source.length) {
    const character = source[index];
    if (character === "'") index = skipQuoted(source, index, "'");
    else if (character === '[') index = skipBrackets(source, index);
    else if (character === '"') break;
    else if (isSpace(character)) {
      const following = trimStart(source.slice(index));
      if (!trimEnd(source.slice(0, index)).endsWith(':') && !following.startsWith(':')) break;
      index = source.length - following.length;
    } else if (separator(character)) break;
    else index += 1;
  }
  return index;
}

function punctuation(source: string, mark: string): number[] {
  const positions: number[] = [];
  let index = 0;
  while (index < source.length) {
    const character = source[index];
    if (character === "'" || character === '"') index = skipQuoted(source, index, character);
    else if (character === '[') index = skipBrackets(source, index);
    else {
      if (character === mark) positions.push(index);
      index += 1;
    }
  }
  return positions;
}

function addressPart(source: string): [string, string] {
  const bangs = punctuation(source, '!');
  if (bangs.length === 0) return ['', source];
  const bang = bangs[bangs.length - 1];
  return [source.slice(0, bang + 1), source.slice(bang + 1)];
}

function parseA1(source: string): Ref | null {
  let index = 0;
  const absCol = source[0] === '$';
  if (absCol) index++;
  const colStart = index;
  while (index < source.length && source[index] >= 'A' && source[index] <= 'Z') index++;
  if (index === colStart) return null;
  let col = 0;
  for (let at = colStart; at < index; at++) {
    col = col * 26 + source.charCodeAt(at) - 64;
    if (col > MAX_COLS) return null;
  }
  const absRow = source[index] === '$';
  if (absRow) index++;
  const rowStart = index;
  while (index < source.length && source[index] >= '0' && source[index] <= '9') index++;
  if (index === rowStart || index !== source.length) return null;
  const row = Number(source.slice(rowStart));
  if (row === 0 || row > MAX_ROWS) return null;
  return { row: row - 1, col: col - 1, absRow, absCol };
}

function column(source: string): Ref | null {
  const letters = source.startsWith('$') ? source.slice(1) : source;
  return /^[A-Za-z]+$/.test(letters) ? parseA1(`${asciiUpper(source)}1`) : null;
}

function row(source: string): Ref | null {
  const digits = source.startsWith('$') ? source.slice(1) : source;
  return /^[0-9]+$/.test(digits) ? parseA1(`A${source}`) : null;
}

function shifted(index: number, absolute: boolean, delta: number, limit: number): number | null {
  const moved = index + (absolute ? 0 : delta);
  return moved >= 0 && moved < limit ? moved : null;
}

/** 0 is A, 26 is AA. */
export function columnLetters(col: number): string {
  let rest = col + 1;
  let letters = '';
  while (rest > 0) {
    rest -= 1;
    letters = String.fromCharCode(65 + (rest % 26)) + letters;
    rest = Math.floor(rest / 26);
  }
  return letters;
}

function toA1(ref: Ref): string {
  return `${ref.absCol ? '$' : ''}${columnLetters(ref.col)}${ref.absRow ? '$' : ''}${ref.row + 1}`;
}

function translateOperand(output: Output, operand: string, rows: number, cols: number, call: boolean): void {
  const parts: string[] = [];
  let start = 0;
  for (const colon of punctuation(operand, ':')) {
    parts.push(operand.slice(start, colon));
    start = colon + 1;
  }
  parts.push(operand.slice(start));
  const threeDimensional = parts.length > 1 && addressPart(parts[0])[0] === '' && addressPart(parts[1])[0] !== '' &&
    parseA1(asciiUpper(trim(parts[0]))) === null;
  const addresses = parts.slice(threeDimensional ? 1 : 0);
  const wholeColumns = addresses.length > 1 && addresses.every((part) => column(trim(addressPart(part)[1])) !== null);
  const wholeRows = addresses.length > 1 && addresses.every((part) => row(trim(addressPart(part)[1])) !== null);
  parts.forEach((part, index) => {
    if (index > 0) output.push(':');
    if ((index === 0 && threeDimensional) || (call && index + 1 === parts.length)) {
      output.push(part);
      return;
    }
    const [prefix, address] = addressPart(part);
    const trimmed = trim(address);
    const cell = wholeColumns ? column(trimmed) : wholeRows ? row(trimmed) : parseA1(asciiUpper(trimmed));
    if (!cell) {
      output.push(part);
      return;
    }
    const movedRow = shifted(cell.row, cell.absRow || wholeColumns, rows, MAX_ROWS);
    const movedCol = shifted(cell.col, cell.absCol || wholeRows, cols, MAX_COLS);
    let replacement = '#REF!';
    if (movedRow !== null && movedCol !== null) {
      if (movedRow === cell.row && movedCol === cell.col) replacement = trimmed;
      else if (wholeColumns) replacement = `${cell.absCol ? '$' : ''}${columnLetters(movedCol)}`;
      else if (wholeRows) replacement = `${cell.absRow ? '$' : ''}${movedRow + 1}`;
      else replacement = toA1({ ...cell, row: movedRow, col: movedCol });
    }
    output.push(prefix);
    output.push(address.slice(0, address.length - trimStart(address).length));
    output.push(replacement);
    output.push(address.slice(trimEnd(address).length));
  });
}

/**
 * A formula, without its `=`, copied from `origin` to `destination`: relative references move by
 * the offset as the engine moves them on paste, and one leaving the sheet becomes `#REF!`.
 */
export function translateFormula(source: string, origin: CellAddr, destination: CellAddr): string {
  if (utf8Length(source) > MAX_FORMULA_BYTES) throw tooLong();
  const rows = destination.row - origin.row;
  const cols = destination.col - origin.col;
  const output = new Output();
  let index = 0;
  while (index < source.length) {
    const start = index;
    const character = source[index];
    if (character === '"') {
      index = skipQuoted(source, index, '"');
      output.push(source.slice(start, index));
    } else if (separator(character)) {
      index += 1;
      output.push(character);
    } else {
      index = operandEnd(source, index);
      translateOperand(output, source.slice(start, index), rows, cols, trimStart(source.slice(index)).startsWith('('));
    }
  }
  return output.text;
}
