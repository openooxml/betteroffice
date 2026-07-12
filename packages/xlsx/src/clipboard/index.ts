/**
 * Pure TSV codec for clipboard interchange, in the Excel dialect (tab-separated
 * fields, CRLF rows, `"`-quoted fields with doubled embedded quotes). No DOM —
 * the chrome hands the system clipboard the string this produces and feeds paste
 * text back through {@link fromTsv}.
 *
 * Security (CLAUDE.md formula-injection contract): a foreign spreadsheet treats
 * a pasted cell as a formula when its text leads with `=`, `+`, `-`, `@`, or a
 * control char that can break the field. We defang such *text* with a leading
 * single quote so it pastes literally — but we export our own real formulas
 * (`isFormula`) verbatim, because faithfully round-tripping a `=SUM(...)` the
 * user actually wrote is the point, and it is our value, not attacker text.
 */

/**
 * One source cell as the wasm range read returns it: the raw entered `input`
 * (a `=...` string when `isFormula`) plus whether the model holds it as a
 * formula.
 */
export interface CellInput {
  input: string;
  isFormula: boolean;
}

const ROW_SEP = '\r\n';
const FIELD_SEP = '\t';

// leading chars a spreadsheet may interpret as a formula on paste.
const INJECTION_LEADERS = new Set(['=', '+', '-', '@']);

// text a foreign spreadsheet could execute or that could break the field: a
// formula leader, or a leading control char (tab/CR/LF).
function needsGuard(text: string): boolean {
  const first = text.charCodeAt(0);
  return INJECTION_LEADERS.has(text[0]) || first === 0x09 || first === 0x0d || first === 0x0a;
}

// quote a field per the excel TSV dialect when it holds a tab, newline, or quote.
function escapeField(text: string): string {
  if (/[\t\n\r"]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function encodeCell(cell: CellInput): string {
  const guarded = !cell.isFormula && cell.input.length > 0 && needsGuard(cell.input);
  return escapeField(guarded ? `'${cell.input}` : cell.input);
}

/**
 * Serialize a row-major grid of cells to a TSV string for the system clipboard,
 * defanging formula-looking text (see the module security note) while exporting
 * genuine formulas as-is.
 */
export function toTsv(cells: CellInput[][]): string {
  return cells.map((row) => row.map(encodeCell).join(FIELD_SEP)).join(ROW_SEP);
}

/**
 * Parse clipboard TSV back into a row-major grid of raw strings. Handles the
 * Excel dialect: `"`-quoted fields with doubled quotes and embedded tabs/
 * newlines, CRLF or LF rows, and a single trailing newline (which does not add
 * an empty row). The formula-injection guard quote is left in place — it is part
 * of the literal text a defanged cell round-trips to.
 */
export function fromTsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;
  let i = 0;

  const pushField = () => {
    row.push(field);
    field = '';
  };
  const pushRow = () => {
    pushField();
    rows.push(row);
    row = [];
  };

  while (i < text.length) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i += 1;
        continue;
      }
      field += ch;
      i += 1;
      continue;
    }
    if (ch === '"') {
      quoted = true;
      i += 1;
      continue;
    }
    if (ch === FIELD_SEP) {
      pushField();
      i += 1;
      continue;
    }
    if (ch === '\r') {
      pushRow();
      // swallow the LF of a CRLF pair.
      if (text[i + 1] === '\n') i += 1;
      i += 1;
      continue;
    }
    if (ch === '\n') {
      pushRow();
      i += 1;
      continue;
    }
    field += ch;
    i += 1;
  }

  // flush the final field/row unless the input ended exactly on a row break.
  if (field.length > 0 || row.length > 0) pushRow();
  return rows;
}
