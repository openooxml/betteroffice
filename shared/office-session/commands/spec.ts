import { hasOwn, type TransferResult } from '../protocol';
import type { CommandTicket, EditSession, SubmitOptions } from './types';

/** @internal */
export type BaseRefusal = 'read-only' | 'stale' | 'superseded' | 'not-ready' | 'busy' | 'disposed'
  | 'invalid-payload' | 'unsupported' | 'dependency-refused' | 'unknown-outcome';
/** @internal */
export const BASE_REFUSALS: readonly BaseRefusal[] = [
  'read-only', 'stale', 'superseded', 'not-ready', 'busy', 'disposed',
  'invalid-payload', 'unsupported', 'dependency-refused', 'unknown-outcome',
];
/** @internal */
export type Refusal<F extends string = never> = { code: BaseRefusal | F; detail?: string };
/** @internal */
export type Dependency = 'doc' | 'selection' | 'layout';
/** @internal Worker state at the read boundary. */
export type At = { incarnation: number; epoch: number; throughSeq: number; doc: number; selection: number; layout: number };
/** @internal */
export type Read<R, F extends string = never> = { ok: true; value: R; at: At } | { ok: false; refusal: Refusal<F> };
/** @internal */
export type Admission = 'view' | 'local-write' | 'remote-update' | 'mode';
declare const payloadType: unique symbol;
declare const queryType: unique symbol;
const REFUSAL = Symbol('session-refusal');

/** @internal Use `satisfies Record<Kind, CommandSpec>` for exhaustive command kinds. */
export interface CommandSpec<P = unknown, F extends string = string> {
  readonly admission: Admission;
  readonly refusals: readonly F[];
  readonly validate?: (payload: unknown) => payload is P;
  readonly [payloadType]?: P;
}
/** @internal */
export type StaleMode = 'answer-at-fence' | { readonly reject: readonly Dependency[] };
/** @internal */
export interface QuerySpec<A extends unknown[] = unknown[], R = unknown, F extends string = string> {
  readonly stale: StaleMode;
  readonly refusals: readonly F[];
  readonly [queryType]?: (...args: A) => R;
}
/** @internal */
export function command<P>() {
  return <const F extends string = never, const X extends object = {}>(
    spec: { admission: Admission; refusals?: readonly F[]; validate?(payload: unknown): payload is P } & X
  ): CommandSpec<P, F> & X => ({ ...spec, refusals: spec.refusals ?? [] });
}
/** @internal */
export function query<A extends unknown[], R>() {
  return <const F extends string = never, const X extends object = {}>(
    spec: { stale: StaleMode; refusals?: readonly F[] } & X
  ): QuerySpec<A, R, F> & X => ({ ...spec, refusals: spec.refusals ?? [] });
}
/** @internal */
export type CommandTable = Record<string, CommandSpec<any, string>>;
/** @internal */
export type QueryTable = Record<string, QuerySpec<any[], any, string>>;
/** @internal */
export type PayloadOf<S> = S extends CommandSpec<infer P, any> ? P : never;
/** @internal */
export type ArgsOf<S> = S extends QuerySpec<infer A, any, any> ? A : never;
/** @internal */
export type ResultOf<S> = S extends QuerySpec<any, infer R, any> ? R : never;
/** @internal */
export type RefusalOf<S> = S extends { readonly refusals: readonly (infer F extends string)[] } ? F : never;
/** @internal */
export type CommandPayload<C extends CommandTable> = PayloadOf<C[keyof C]>;
/** @internal */
export interface RefusalResult<F extends string> {
  readonly [REFUSAL]: true;
  readonly refusal: Refusal<F>;
}
/** @internal */
export type QueryResult<R, F extends string> = R | TransferResult<R> | RefusalResult<F>;
/** @internal `satisfies QueryTable` preserves kinds for exhaustive `QueryHandlers<typeof table>`. */
export type QueryHandlers<Q extends QueryTable> = {
  [K in keyof Q]: (args: string extends keyof Q ? any : ArgsOf<Q[K]>, context: unknown) =>
    QueryResult<ResultOf<Q[K]>, RefusalOf<Q[K]>> | Promise<QueryResult<ResultOf<Q[K]>, RefusalOf<Q[K]>>>;
};
/** @internal */
export function refuse<const F extends string>(code: BaseRefusal | F, detail?: string): RefusalResult<F> {
  return { [REFUSAL]: true, refusal: { code, ...(detail === undefined ? {} : { detail }) } };
}
/** @internal */
export function isRefusalResult(value: unknown): value is RefusalResult<string> {
  return record(value) && value[REFUSAL] === true;
}
/** @internal Edge validation for unknown kinds and payloads. */
export function checkCommand<C extends CommandTable>(table: C, kind: string, payload: unknown): Refusal | undefined {
  if (!hasOwn(table, kind)) return { code: 'unsupported' };
  return table[kind]!.validate && !table[kind]!.validate!(payload) ? { code: 'invalid-payload' } : undefined;
}
/** @internal */
export type QueryFence = { session: string; incarnation: number; epoch: number; afterSeq: number };
/** @internal */
export type QueryReply = Read<unknown, string>;
/** @internal */
export type QueryMethods = { query(fence: QueryFence, kind: string, args: unknown[]): QueryReply };
/** @internal */
export function isQueryReply(value: unknown): value is QueryReply {
  if (!record(value)) return false;
  if (value.ok === false) return record(value.refusal) && typeof value.refusal.code === 'string' &&
    (value.refusal.detail === undefined || typeof value.refusal.detail === 'string');
  return value.ok === true && hasOwn(value, 'value') && record(value.at) &&
    ['incarnation', 'epoch', 'throughSeq', 'doc', 'selection', 'layout'].every((key) => counter(value.at[key]));
}
/** @internal */
export type HoldReason = 'composition' | 'draft' | 'gesture' | 'mode';
/** @internal */
export interface Hold { readonly reason: HoldReason; readonly released: boolean; release(): void }
/** @internal */
export type ReadOptions = { within?: Hold; signal?: AbortSignal };
/** @internal */
export type SessionStage = 'ready' | 'recovering' | 'failed' | 'disposed';
/** @internal */
export interface SessionReads<Q extends QueryTable> {
  read<K extends keyof Q & string>(kind: K, args: ArgsOf<Q[K]>, options?: ReadOptions):
    Promise<Read<ResultOf<Q[K]>, RefusalOf<Q[K]>>>;
  hold(reason: HoldReason, options?: { signal?: AbortSignal }): Hold;
  readonly stage: SessionStage;
}
/** @internal */
export type WorkerEditSession<P, S, E, X, Q extends QueryTable = {}> = EditSession<P, S, E, X> & SessionReads<Q>;
/** @internal Per-kind payload view of the same session object. */
export type CommandSession<C extends CommandTable, S, E, X, Q extends QueryTable = {}> =
  Omit<WorkerEditSession<CommandPayload<C>, S, E, X, Q>, 'submit'> & {
    submit<K extends keyof C & string>(kind: K, payload: PayloadOf<C[K]>, options?: SubmitOptions<X>): CommandTicket;
  };

function record(value: unknown): value is Record<PropertyKey, any> {
  return value !== null && typeof value === 'object';
}

function counter(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}
