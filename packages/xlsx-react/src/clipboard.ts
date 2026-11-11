import { fromTsv, toTsv } from '@betteroffice/xlsx';
import type { CellInputEdit, CellRange } from '@betteroffice/xlsx';
import type { TranslationKey } from '@betteroffice/xlsx-i18n';
import { WorkerInputRefusal } from './commands/workerInputErrors';

const MAX_CLIPBOARD_CELLS = 100_000;
const MAX_CLIPBOARD_BYTES = 16 * 1024 * 1024;

export type ClipboardMethod = 'read' | 'readText' | 'write' | 'writeText';
/** Why the page cannot use the clipboard: no secure origin, no API, or the browser refused it. */
export type ClipboardRefusalCause = 'insecure' | 'unsupported' | 'denied';

const REFUSAL_MESSAGES: Record<ClipboardRefusalCause, string> = {
  insecure: 'The clipboard needs a secure (HTTPS) page',
  unsupported: 'The clipboard is unavailable in this browser',
  denied: 'The browser did not allow access to the clipboard',
};

const WRITE_MESSAGES: Record<ClipboardRefusalCause, string> = {
  insecure: REFUSAL_MESSAGES.insecure,
  unsupported: 'This browser does not allow copying to the clipboard',
  denied: REFUSAL_MESSAGES.denied,
};

const REFUSAL_KEYS: Record<ClipboardRefusalCause, TranslationKey> = {
  insecure: 'commands.reasons.clipboardInsecure',
  unsupported: 'commands.reasons.clipboardUnavailable',
  denied: 'commands.reasons.clipboardDenied',
};

const WRITE_KEYS: Record<ClipboardRefusalCause, TranslationKey> = {
  insecure: 'commands.reasons.clipboardInsecure',
  unsupported: 'commands.reasons.clipboardWriteUnavailable',
  denied: 'commands.reasons.clipboardDenied',
};

/** A clipboard call the page cannot make. Never means a pending edit failed. */
export class ClipboardRefusalError extends WorkerInputRefusal {
  constructor(readonly unavailable: ClipboardRefusalCause, readonly method: ClipboardMethod, barrier = false) {
    super((method.startsWith('write') ? WRITE_MESSAGES : REFUSAL_MESSAGES)[unavailable], barrier,
      unavailable === 'denied' ? 'clipboard-denied' : 'clipboard-unavailable');
    this.name = 'ClipboardRefusalError';
  }
}

/** Why the page cannot use a clipboard method, or `null` when it can. */
export function clipboardUnavailable(method: ClipboardMethod): ClipboardRefusalCause | null {
  if (typeof window !== 'undefined' && window.isSecureContext === false) return 'insecure';
  const clipboard: Partial<Record<ClipboardMethod, unknown>> | undefined =
    typeof navigator === 'undefined' ? undefined : navigator.clipboard;
  return typeof clipboard?.[method] === 'function' ? null : 'unsupported';
}

/**
 * Starts a clipboard call inside the gesture that authorizes it, so the browser keeps the user
 * activation. A missing method is reported before the call, and a denial or an aborted call from
 * the promise becomes a {@link ClipboardRefusalError}; other failures keep their own error.
 */
export function guardedClipboard<T>(
  method: ClipboardMethod, call: () => Promise<T>
): { call: Promise<T> } | { unavailable: ClipboardRefusalCause } {
  const unavailable = clipboardUnavailable(method);
  if (unavailable) return { unavailable };
  try {
    return { call: call().catch((error) => { throw clipboardFailure(error, method) ?? error; }) };
  } catch {
    return { unavailable: 'denied' };
  }
}

function clipboardFailure(error: unknown, method: ClipboardMethod): ClipboardRefusalError | null {
  if (!error || typeof error !== 'object' || !('name' in error)) return null;
  if (error.name === 'AbortError') {
    // The OS clipboard aborted the call (another app owns it, or the browser withdrew the write).
    // The page cannot complete the operation, but the workbook edit behind it was not refused.
    return new ClipboardRefusalError('unsupported', method);
  }
  return error.name === 'NotAllowedError' || error.name === 'SecurityError'
    ? new ClipboardRefusalError('denied', method) : null;
}

/** The localized message for a clipboard refusal; `null` for any other error. */
export function clipboardRefusalMessage(
  error: unknown, translate: (key: TranslationKey) => string
): string | null {
  return error instanceof ClipboardRefusalError
    ? translate((error.method.startsWith('write') ? WRITE_KEYS : REFUSAL_KEYS)[error.unavailable]) : null;
}

type ClipboardCells = { input: string; isFormula: boolean; isArrayFormula?: boolean }[][];
type Capture = { owner: object; sheet: number; range: CellRange; cells: ClipboardCells; cut: boolean };
type Token = { revision: number; cancellation: number; escapes: number; marker: string };
export type CellCapture = Capture & { text: string; marker: string };
export type CellPaste = Capture & { edits: CellInputEdit[]; bottom: number; right: number };

export class CellClipboard {
  pending: Promise<unknown> = Promise.resolve();
  private revision = 0;
  private cancellation = 0;
  private escapes = 0;
  private watching: Token | null = null;
  private captured: CellCapture | null = null;

  get escaped(): number {
    return this.escapes;
  }

  get current(): CellCapture | null {
    return this.captured;
  }

  begin(escapes = this.escapes, marker = newClipboardMarker()): Token {
    this.captured = null;
    this.watching = null;
    return { revision: ++this.revision, cancellation: this.cancellation, escapes, marker };
  }

  capture(token: Token, capture: Capture, marked = false): void {
    if (!marked || token.revision !== this.revision ||
      capture.cut && (token.cancellation !== this.cancellation || token.escapes !== this.escapes)) return;
    const { owner, ...data } = capture;
    this.captured = { ...structuredClone(data), owner, text: toTsv(capture.cells), marker: token.marker };
  }

  watch(token: Token): void {
    if (token.revision === this.revision) this.watching = token;
  }

  changed(): void {
    if (this.captured?.cut || this.watching) this.cancelCut();
  }

  escape(): void {
    this.escapes++;
    this.cancelCut();
  }

  cancelCut(): void {
    this.cancellation++;
    this.watching = null;
    if (this.captured?.cut) this.captured = null;
  }

  paste(
    owner: object, text: string, target: CellRange, html?: string | null, sheet?: number, accepted?: CellCapture | null
  ): CellPaste | null {
    const matches = (capture: CellCapture | null | undefined): capture is CellCapture => Boolean(capture &&
      capture.owner === owner && capture.text === text && capture.marker && html?.includes(capture.marker));
    const source = [accepted?.cut ? null : accepted, this.captured].find(matches);
    if (!source) {
      this.cancelCut();
      return null;
    }
    if (source.cut && sheet !== undefined && source.sheet !== sheet) {
      this.cancelCut();
      throw new WorkerInputRefusal('Moving cells between worksheets is unavailable');
    }
    if (!source.cut && source.cells.some((row) => row.some((cell) => cell.isArrayFormula))) {
      throw new WorkerInputRefusal('Copying array or spill formula cells is unavailable');
    }
    const rows = source.range.bottom - source.range.top + 1;
    const cols = source.range.right - source.range.left + 1;
    const height = source.cut ? rows : Math.max(rows, target.bottom - target.top + 1);
    const width = source.cut ? cols : Math.max(cols, target.right - target.left + 1);
    if (height * width > MAX_CLIPBOARD_CELLS) throw new WorkerInputRefusal('Paste exceeds the 100000-cell limit');
    const inputs = parseCellClipboardText(text);
    const edits: CellInputEdit[] = [];
    for (let dr = 0; dr < height; dr++) {
      for (let dc = 0; dc < width; dc++) {
        const row = dr % rows;
        const col = dc % cols;
        const input = inputs[row]?.[col] ?? '';
        edits.push({ row: target.top + dr, col: target.left + dc, input,
          ...(source.cells[row]?.[col]?.isFormula ? {
            copyFrom: { row: source.range.top + row, col: source.range.left + col },
          } : {}),
        });
      }
    }
    return { ...source, edits, bottom: target.top + height - 1, right: target.left + width - 1 };
  }

  consume(paste: CellPaste, applied: boolean): void {
    if (applied && paste.cut && this.captured?.owner === paste.owner && this.captured.cells === paste.cells) this.cancelCut();
  }
}

export function newClipboardMarker(): string {
  return `betteroffice-${crypto.randomUUID()}`;
}

export function clipboardMarker(token: Pick<Token, 'marker'>, text = ''): Blob {
  const escape = (value: string) => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const rows = fromTsv(text).map((row) => `<tr>${row.map((cell) => `<td>${escape(cell)}</td>`).join('')}</tr>`).join('');
  return new Blob([`<table data-betteroffice-clipboard="${token.marker}">${rows}</table>`], { type: 'text/html' });
}

export function clipboardHtml(token: Pick<Token, 'marker'>, text: Promise<Blob>): Promise<Blob> {
  const html = text.then(async (blob) => clipboardMarker(token, await blob.text()));
  void html.catch(() => {});
  return html;
}

function validateClipboardText(text: string): string {
  if (text.length > MAX_CLIPBOARD_BYTES || new TextEncoder().encode(text).byteLength > MAX_CLIPBOARD_BYTES) {
    throw new WorkerInputRefusal('Clipboard exceeds the 16 MiB limit');
  }
  return text;
}

export function parseCellClipboardText(text: string): string[][] {
  validateClipboardText(text);
  let quoted = false;
  let fields = 0;
  for (let at = 0; at < text.length; at++) {
    const ch = text[at];
    if (ch === '"') {
      if (quoted && text[at + 1] === '"') at++;
      else quoted = !quoted;
    } else if (!quoted && (ch === '\t' || ch === '\r' || ch === '\n')) {
      if (++fields > MAX_CLIPBOARD_CELLS) throw new WorkerInputRefusal('Paste exceeds the 100000-cell limit');
      if (ch === '\r' && text[at + 1] === '\n') at++;
    }
  }
  const rows = fromTsv(text);
  if (rows.reduce((count, row) => count + row.length, 0) > MAX_CLIPBOARD_CELLS) {
    throw new WorkerInputRefusal('Paste exceeds the 100000-cell limit');
  }
  return rows;
}

export function writeClipboardText(text: string): Promise<void> {
  const access = guardedClipboard('writeText', () => navigator.clipboard.writeText(text));
  return 'call' in access ? access.call : Promise.reject(new ClipboardRefusalError(access.unavailable, 'writeText'));
}

/** Starts a rich clipboard write in the gesture; the item is built inside it so a throw there is a refusal. */
export function writeClipboard(item: () => ClipboardItem): Promise<void> {
  const access = guardedClipboard('write', () => navigator.clipboard.write([item()]));
  return 'call' in access ? access.call : Promise.reject(new ClipboardRefusalError(access.unavailable, 'write'));
}

/** Plain text only; refuses with {@link ClipboardRefusalError} when the page cannot read it. */
export async function readClipboardText(): Promise<string> {
  const access = guardedClipboard('readText', () => navigator.clipboard.readText());
  if ('unavailable' in access) throw new ClipboardRefusalError(access.unavailable, 'readText');
  return validateClipboardText(await access.call);
}

export async function readCellClipboard(): Promise<{ text: string; html?: string | null }> {
  const access = guardedClipboard('readText', () => navigator.clipboard.readText());
  if ('unavailable' in access) throw new ClipboardRefusalError(access.unavailable, 'readText');
  const reading = access.call.then(validateClipboardText);
  void reading.catch(() => {});
  const rich = guardedClipboard('read', () => navigator.clipboard.read());
  if ('unavailable' in rich) return { text: await reading };
  let blob: Blob | null = null;
  try {
    const items = await rich.call;
    const item = items.find((item) => item.types.includes('text/plain'));
    if (item?.types.includes('text/html')) blob = await item.getType('text/html');
  } catch {}
  if (blob && blob.size > MAX_CLIPBOARD_BYTES) throw new WorkerInputRefusal('Clipboard exceeds the 16 MiB limit');
  const html = blob ? await blob.text() : null;
  return { text: await reading, html };
}
