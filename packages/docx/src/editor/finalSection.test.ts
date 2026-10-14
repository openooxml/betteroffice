import { describe, expect, test } from 'bun:test';

import type { DocumentBody, SectionProperties } from '../types/document';
import { resolvedFinalSectionProperties, updateFinalSectionProperties } from './finalSection';

const inherited: NonNullable<SectionProperties['headerReferences']> = [
  { type: 'default', rId: 'rId3' },
];
const body: DocumentBody = {
  content: [],
  sections: [
    { properties: { headerReferences: inherited, titlePg: true }, content: [] },
    { properties: { headerReferences: inherited, titlePg: true, marginLeft: 1440 }, content: [] },
  ],
  finalSectionProperties: { marginLeft: 1440 },
};

describe('resolvedFinalSectionProperties', () => {
  test('reads the last section entry, which carries the inherited references', () => {
    expect(resolvedFinalSectionProperties(body)).toBe(body.sections![1].properties);
  });

  test('falls back to the authored body sectPr when sections are absent', () => {
    const authored = { marginLeft: 720 };
    expect(
      resolvedFinalSectionProperties({ content: [], finalSectionProperties: authored })
    ).toBe(authored);
    expect(resolvedFinalSectionProperties(null)).toBeUndefined();
  });
});

describe('updateFinalSectionProperties', () => {
  test('edits the resolved entry and the authored sectPr alike, leaving earlier sections alone', () => {
    const next = updateFinalSectionProperties(body, (properties) => ({
      ...properties,
      marginLeft: 2880,
    }));
    expect(next.sections![1].properties).toEqual({
      headerReferences: inherited,
      titlePg: true,
      marginLeft: 2880,
    });
    expect(next.finalSectionProperties).toEqual({ marginLeft: 2880 });
    expect(next.sections![0]).toBe(body.sections![0]);
    expect(body.sections![1].properties.marginLeft).toBe(1440);
  });

  test('creates the body sectPr when none was authored', () => {
    const next = updateFinalSectionProperties({ content: [] }, (properties) => ({
      ...properties,
      titlePg: true,
    }));
    expect(next.finalSectionProperties).toEqual({ titlePg: true });
    expect(next.sections).toBeUndefined();
  });
});
