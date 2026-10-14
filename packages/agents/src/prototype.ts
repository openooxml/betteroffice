import { DocumentToolError, type GrepOptions, type RenderedPage } from './types';

export interface PrototypeOptions {
  name?: string;
  wasm?: Uint8Array;
}

export interface AgentProposal<Change> {
  id: string;
  author: string;
  note: string;
  status: 'pending' | 'accepted' | 'rejected';
  version: string;
  changes: Change[];
  stale: boolean;
}

export function integer(value: number, min: number, max: number, name: string) {
  if (!Number.isSafeInteger(value) || value < min || value > max) {
    throw new DocumentToolError('INVALID_ARGUMENT', `${name} must be an integer from ${min} to ${max}.`);
  }
  return value;
}

export function unwrap<T extends { ok: boolean }>(value: T): Extract<T, { ok: true }> {
  if (!value.ok) {
    const { failure } = value as T & { failure: { code: string; message: string } };
    throw new DocumentToolError(failure.code, failure.message, failure);
  }
  return value as Extract<T, { ok: true }>;
}

export function textWindow(text: string, options: { start?: number; length?: number } = {}) {
  let start = integer(options.start ?? 0, 0, text.length, 'start');
  const length = integer(options.length ?? 4000, 1, 16000, 'length');
  const split = (i: number) => /[\uD800-\uDBFF]/.test(text.charAt(i - 1)) && /[\uDC00-\uDFFF]/.test(text.charAt(i));
  if (split(start)) start--;
  let end = Math.min(text.length, start + length);
  if (split(end)) end--;
  if (end === start && end < text.length) end += 2;
  return { text: text.slice(start, end), start, end, length: text.length, nextStart: end < text.length ? end : null };
}

export function plainText(value: string) {
  if (typeof value !== 'string' || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value)) {
    throw new DocumentToolError('INVALID_ARGUMENT', 'Text must contain complete Unicode characters and no XML control characters.');
  }
}

export async function initialize(format: 'xlsx' | 'pptx', init: (input?: Uint8Array) => Promise<void>, wasm?: Uint8Array) {
  if (!wasm && typeof process !== 'undefined' && process.versions?.node) {
    const { readFile } = await import('node:fs/promises');
    const entry = format === 'xlsx' ? import.meta.resolve('@betteroffice/xlsx') : import.meta.resolve('@betteroffice/pptx');
    const directory = entry.endsWith('/src/index.ts') ? './wasm/generated' : './generated';
    wasm = await readFile(new URL(`${directory}/${format}_wasm_bg.wasm`, entry));
  }
  await init(wasm);
}

export abstract class PrototypeDocument<Change, Step> {
  private readonly proposals = new Map<string, { value: AgentProposal<Change>; steps: Step[] }>();
  private readonly cursors = new Map<string, { signature: string; version: string; offset: number }>();
  private nextId = 1;
  protected readonly prefix = crypto.randomUUID();
  protected closed = false;
  abstract readonly format: 'xlsx' | 'pptx';
  abstract readonly name: string;
  protected abstract currentVersion(): string;
  protected abstract validate(steps: Step[], version: string): void;
  protected abstract apply(steps: Step[], version: string): unknown;
  protected abstract save(steps?: Step[]): Uint8Array;
  protected abstract reopen(bytes: Uint8Array): void;
  protected abstract dispose(): void;

  protected alive() {
    if (this.closed) throw new DocumentToolError('CLOSED', 'This document is closed. Open it again.');
  }

  protected stage(input: { author: string; note?: string }, changes: Change[], steps: Step[], version: string) {
    this.alive();
    if (!input.author || input.author.length > 200 || (input.note?.length ?? 0) > 2000 || !changes.length || changes.length > 32) {
      throw new DocumentToolError('INVALID_ARGUMENT', 'Supply an author (1–200 characters), note (up to 2000), and 1–32 edits.');
    }
    if (JSON.stringify(changes).length > 32000) throw new DocumentToolError('EDIT_LIMIT', 'Combined proposal changes must fit in 32000 UTF-16 units.');
    this.validate(steps, version);
    if (this.proposals.size >= 64) {
      const old = [...this.proposals].find(([, held]) => held.value.status !== 'pending');
      if (old) this.proposals.delete(old[0]);
      else throw new DocumentToolError('PROPOSAL_LIMIT', 'Reject or discard a proposal before staging more than 64.');
    }
    const value: AgentProposal<Change> = { id: `proposal${this.nextId++}`, author: input.author, note: input.note ?? '', status: 'pending', version, changes, stale: false };
    this.proposals.set(value.id, { value: structuredClone(value), steps: structuredClone(steps) });
    return value;
  }

  review(id: string) {
    const { value } = this.held(id);
    return structuredClone({ ...value, stale: value.status === 'pending' && value.version !== this.currentVersion() });
  }

  listProposals() {
    this.alive();
    return [...this.proposals.values()].map(({ value }) => ({ id: value.id, author: value.author, note: value.note.slice(0, 120), status: value.status, changes: value.changes.length, stale: value.status === 'pending' && value.version !== this.currentVersion() }));
  }

  async accept(id: string, options: { tracked?: boolean; date?: string } = {}) {
    if (options.tracked || options.date) throw new DocumentToolError('UNSUPPORTED', 'Tracked acceptance is available for DOCX only. Omit tracked and date.');
    const held = this.pending(id);
    const receipt = this.apply(held.steps, held.value.version);
    held.value.status = 'accepted';
    return { id, status: 'accepted' as const, receipt };
  }

  reject(id: string) {
    const held = this.held(id);
    if (held.value.status !== 'pending') throw new DocumentToolError('PROPOSAL_CLOSED', `Proposal is ${held.value.status}.`);
    held.value.status = 'rejected';
    return { id, status: 'rejected' as const };
  }

  discard(id: string) { this.held(id); this.proposals.delete(id); }

  async export(id?: string) {
    this.alive();
    return this.save(id ? this.pending(id).steps : undefined);
  }

  async verify(id?: string) {
    const bytes = await this.export(id);
    this.reopen(bytes);
    return { reopened: true, bytes: bytes.length, checks: ['save', 'reopen'], unchecked: ['visual fidelity', 'semantic correctness', ...(this.format === 'xlsx' ? ['formula correctness'] : [])] };
  }

  async render(_page?: number, _proposal?: string): Promise<RenderedPage> {
    this.alive();
    throw new DocumentToolError('UNSUPPORTED', `PNG rendering is not available in the ${this.format.toUpperCase()} prototype. Review the proposal changes and verify its export.`);
  }

  close() {
    if (this.closed) return;
    this.closed = true;
    this.proposals.clear();
    this.cursors.clear();
    this.dispose();
  }

  protected searchPage<T>(options: GrepOptions, hits: T[], clipped = false) {
    const limit = integer(options.limit ?? 20, 1, 100, 'limit');
    const signature = JSON.stringify([options.query, options.caseSensitive, options.story]);
    const version = this.currentVersion();
    const cursor = options.cursor ? this.cursors.get(options.cursor) : undefined;
    if (options.cursor && (!cursor || cursor.version !== version || cursor.signature !== signature)) {
      throw new DocumentToolError('STALE_CURSOR', 'Search changed or the document was edited. Repeat grep without cursor.');
    }
    const offset = cursor?.offset ?? 0;
    const matches: T[] = [];
    let size = 0;
    for (const hit of hits.slice(offset, offset + limit)) {
      const cost = JSON.stringify(hit).length + 100;
      if (matches.length && size + cost > 16000) break;
      matches.push(hit); size += cost;
    }
    let nextCursor: string | null = null;
    if (offset + matches.length < hits.length) {
      nextCursor = `${this.prefix}:cursor${this.nextId++}`;
      if (this.cursors.size >= 64) this.cursors.delete(this.cursors.keys().next().value!);
      this.cursors.set(nextCursor, { signature, version, offset: offset + matches.length });
    }
    return { matches, nextCursor, version, truncated: clipped, ...(clipped ? { warning: 'Search stopped at 10000 matches. Narrow query or story to search beyond this cap.' } : {}) };
  }

  protected checkQuery(options: GrepOptions) {
    this.alive();
    if (!options.query || options.query.length > 1000 || /[\r\n]/.test(options.query)) throw new DocumentToolError('INVALID_ARGUMENT', 'query must be 1–1000 characters within one paragraph or cell.');
  }

  private held(id: string) {
    this.alive();
    const held = this.proposals.get(id);
    if (!held) throw new DocumentToolError('UNKNOWN_PROPOSAL', 'Stage a proposal first.');
    return held;
  }

  private pending(id: string) {
    const held = this.held(id);
    if (held.value.status !== 'pending') throw new DocumentToolError('PROPOSAL_CLOSED', `Proposal is ${held.value.status}.`);
    if (held.value.version !== this.currentVersion()) throw new DocumentToolError('STALE_TARGET', 'The document changed. Read fresh targets and create a new proposal.');
    return held;
  }
}
