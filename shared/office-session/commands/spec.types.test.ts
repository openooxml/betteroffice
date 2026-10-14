import { expect, it } from 'bun:test';
import {
  command, query, refuse, checkCommand, isQueryReply,
  type ArgsOf, type CommandSession, type CommandSpec, type CommandTable,
  type PayloadOf, type QueryHandlers, type QueryTable, type RefusalOf, type ResultOf,
} from './spec';

const commands = {
  insert: command<{ text: string }>()({ admission: 'local-write', refusals: ['empty'], authored: true }),
  select: command<number>()({ admission: 'view' }),
} satisfies CommandTable;
const queries = {
  text: query<[start: number], string>()({ stale: { reject: ['doc'] }, refusals: ['range'] }),
} satisfies QueryTable;

function types(session: CommandSession<typeof commands, number, boolean, never, typeof queries>) {
  const payload: PayloadOf<typeof commands.insert> = { text: 'a' };
  const args: ArgsOf<typeof queries.text> = [0];
  const result: ResultOf<typeof queries.text> = 'a';
  const refusal: RefusalOf<typeof queries.text> = 'range';
  const authored: true = commands.insert.authored;
  session.submit('insert', payload);
  session.read('text', args);
  // @ts-expect-error Wrong per-kind payload.
  session.submit('insert', 1);
  // @ts-expect-error Unknown command kind.
  session.submit('missing', payload);
  // @ts-expect-error Wrong query arguments.
  session.read('text', ['a']);
  // @ts-expect-error Missing query handler.
  const missing: QueryHandlers<typeof queries> = {};
  // @ts-expect-error Undeclared query refusal.
  const handlers: QueryHandlers<typeof queries> = { text: () => refuse('other') };
  // @ts-expect-error Missing command kind.
  const table = { insert: commands.insert } satisfies Record<'insert' | 'select', CommandSpec>;
  // @ts-expect-error Refusal inference preserves literals.
  const wrong: RefusalOf<typeof queries.text> = 'other';
  return { result, refusal, authored, missing, handlers, table, wrong };
}
void types;

it('defaults refusal lists and validates command edges and read stamps', () => {
  expect(commands.select.refusals).toEqual([]);
  expect(checkCommand(commands, 'missing', null)).toEqual({ code: 'unsupported' });
  const validated = { insert: command<number>()({ admission: 'local-write',
    validate: (value): value is number => typeof value === 'number' }) };
  expect(checkCommand(validated, 'insert', 'a')).toEqual({ code: 'invalid-payload' });
  expect(checkCommand(validated, 'insert', 1)).toBeUndefined();
  const at = { incarnation: 0, epoch: 1, throughSeq: 0, doc: 0, selection: 0, layout: 0 };
  expect(isQueryReply({ ok: true, value: null, at })).toBe(true);
  expect(isQueryReply({ ok: true, value: null, at: { ...at, layout: -1 } })).toBe(false);
  expect(isQueryReply({ ok: false, refusal: { code: 'busy' } })).toBe(true);
  expect(isQueryReply({ ok: false, refusal: { code: 'busy', detail: 1 } })).toBe(false);
});
