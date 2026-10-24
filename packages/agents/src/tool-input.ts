import { z } from 'zod';

export const dateSchema = z.union([z.string().datetime(), z.iso.date()])
  .transform(value => value.length === 10 ? `${value}T00:00:00Z` : value)
  .describe('ISO timestamp or YYYY-MM-DD, normalized to midnight UTC for date-only input.');

export function readLimit(cap: number) {
  return z.number().int().min(1).transform(value => Math.min(value, cap)).optional()
    .describe(`Maximum ${cap}; larger requests are clamped. Follow pagination cursors for remaining content.`);
}

export function clampingNotes(input: unknown, parsed: unknown) {
  const original = record(input);
  return Object.entries(record(parsed)).flatMap(([key, value]) =>
    ['limit', 'shapeLimit', 'sectionLimit', 'textLength', 'length'].includes(key) && typeof value === 'number' && typeof original[key] === 'number' && original[key] > value
      ? [`${key} clamped to ${value}; follow pagination cursors for remaining content.`] : []);
}

type Schema = { properties?: Record<string, Schema>; items?: Schema; oneOf?: Schema[]; anyOf?: Schema[]; required?: string[]; const?: unknown };
function record(value: unknown): Record<PropertyKey, unknown> {
  return value !== null && typeof value === 'object' ? value as Record<PropertyKey, unknown> : {};
}
function inputAt(input: unknown, path: PropertyKey[]) {
  return path.reduce<unknown>((value, key) => record(value)[key], input);
}
function schemaAt(schema: unknown, path: PropertyKey[]) {
  let node = schema as Schema | undefined;
  for (const key of path) node = typeof key === 'number' ? node?.items : node?.properties?.[String(key)];
  return node;
}
function distance(a: string, b: string) {
  let row = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let i = 1; i <= a.length; i++) {
    const next = [i];
    for (let j = 1; j <= b.length; j++) next[j] = Math.min(row[j] + 1, next[j - 1] + 1, row[j - 1] + Number(a[i - 1] !== b[j - 1]));
    row = next;
  }
  return row[b.length];
}
function signature(schema: Schema, key: string) {
  const required = (schema.required ?? []).filter(field => field !== key);
  const optional = Object.keys(schema.properties ?? {}).filter(field => field !== key && !required.includes(field)).map(field => `${field}?`);
  return `${schema.properties?.[key]?.const} {${[...required, ...optional].join(', ')}}`;
}

export function toolInputIssues(error: z.ZodError, schema: unknown, input: unknown) {
  return error.issues.map(issue => {
    let message = issue.message;
    const key = issue.path.at(-1);
    if (issue.code === 'invalid_union' && (key === 'op' || key === 'action')) {
      const parent = schemaAt(schema, issue.path.slice(0, -1));
      const variants = (parent?.oneOf ?? parent?.anyOf ?? []).filter(variant => typeof variant.properties?.[key]?.const === 'string');
      const target = record(inputAt(input, issue.path.slice(0, -1)));
      const value = target[key] ?? target.type;
      const attempted = typeof value === 'string' ? value.slice(0, 128) : '';
      const score = (name: string) => distance(attempted, name) - (attempted.split('_').at(-1) === name.split('_').at(-1) ? attempted.length : 0);
      const closest = variants.slice().sort((a, b) => score(String(a.properties![key].const)) - score(String(b.properties![key].const)))[0];
      if (closest) message = `Invalid ${key} ${JSON.stringify(attempted)}. Valid ${key === 'op' ? 'ops' : 'actions'}: ${variants.map(variant => variant.properties![key].const).join(', ')}. Closest: ${signature(closest, key)}. Use the literal ${key} and field names.`;
    }
    const given = inputAt(input, issue.path);
    const op = record(inputAt(input, issue.path.slice(0, -1))).op;
    if (issue.code === 'invalid_type' && typeof given === 'number' && (key === 'slide' || ((key === 'after' || key === 'before') && (op === 'move_slide' || op === 'duplicate_slide')))) {
      message = `Slide IDs are strings. ${given} could be a 0-based index or a 1-based slide number, so it is refused. Copy a slide ID such as slide:<0-based index>:<source ID> from pptx_outline.`;
    }
    if (issue.code === 'unrecognized_keys' && issue.keys.includes('mapping') && record(inputAt(input, issue.path)).op === 'set_slide_layout') {
      message += ' Use placeholders:[{shape, layoutPlaceholder}] for set_slide_layout; copy shape IDs and target layout placeholder IDs from pptx_outline.';
    }
    return { path: issue.path, message };
  });
}
