import { describe, expect, it } from 'bun:test';

import type { Paragraph } from '../types/content/paragraph';
import { canonicalProjection, checkSourceOrdinals } from './projectionParity.support';

describe('DOCX projection canonical form', () => {
  it('checks paragraph source ordinals against the source pilcrow key', () => {
    const paragraph: Paragraph = { type: 'paragraph', content: [] };
    const source = { ...paragraph, sourceOrdinal: 3 };
    const edited = { ...paragraph, content: [{ type: 'run', content: [] }] } satisfies Paragraph;
    expect(canonicalProjection(source)).toEqual(canonicalProjection(paragraph));
    expect(canonicalProjection(source)).not.toEqual(canonicalProjection(edited));
    expect(checkSourceOrdinals('body', [edited, source], [source], ['new', 'body:p0'])).toEqual([]);
    expect(checkSourceOrdinals('body', [{ ...source, sourceOrdinal: 4 }], [source], ['body:p0']))
      .toHaveLength(1);
    expect(canonicalProjection({ metadata: { sourceOrdinal: 3 } })).not
      .toEqual(canonicalProjection({ metadata: { sourceOrdinal: 4 } }));
  });

  it('equates a redundant whitespace hint while keeping a significant one', () => {
    const text = { type: 'text', text: ' some text' };
    const plain = { type: 'text', text: 'some text' };
    expect(canonicalProjection({ ...text, preserveSpace: true })).toEqual(canonicalProjection(text));
    expect(canonicalProjection({ ...plain, preserveSpace: true })).not.toEqual(canonicalProjection(plain));
  });

  it('equates empty run formatting with absence and preserves explicit formatting', () => {
    const run = { type: 'run', content: [] };
    expect(canonicalProjection({ ...run, formatting: {} })).toEqual(canonicalProjection(run));
    expect(canonicalProjection({ ...run, formatting: { bold: false } })).not.toEqual(canonicalProjection(run));
  });

  it('equates an empty table cascade with absence and preserves a selected style', () => {
    const table = { type: 'table', rows: [] };
    expect(canonicalProjection({ ...table, formatting: { styleCascade: {} } }))
      .toEqual(canonicalProjection(table));
    expect(canonicalProjection({ ...table, formatting: { styleCascade: { selectedStyleId: 'TableGrid' } } }))
      .not.toEqual(canonicalProjection(table));
  });

  it('equates omitted field display mode with result and preserves code mode', () => {
    const field = { type: 'simpleField', fieldTree: { version: 1 } };
    expect(canonicalProjection({ ...field, fieldTree: { ...field.fieldTree, displayMode: 'result' } }))
      .toEqual(canonicalProjection(field));
    expect(canonicalProjection({ ...field, fieldTree: { ...field.fieldTree, displayMode: 'code' } }))
      .not.toEqual(canonicalProjection(field));
  });

  it('equates absent and empty separator collections and preserves separator notes', () => {
    const pkg = { document: { content: [] } };
    for (const key of ['footnoteSeparators', 'endnoteSeparators']) {
      expect(canonicalProjection({ package: { ...pkg, [key]: [] } }))
        .toEqual(canonicalProjection({ package: pkg }));
      expect(canonicalProjection({ package: { ...pkg, [key]: [{ id: -1, content: [] }] } }))
        .not.toEqual(canonicalProjection({ package: pkg }));
    }
  });
});
