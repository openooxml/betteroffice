import { describe, expect, test } from 'bun:test';
import { buildResidentRegionLayoutRequest } from './computeLayout';
import type { Document } from '../types/document';

function documentWith(body: Record<string, unknown>): Document {
  return {
    package: {
      document: body,
      footnotes: [],
      endnotes: [],
    },
  } as unknown as Document;
}

describe('buildResidentRegionLayoutRequest sections', () => {
  test('a single-section document yields exactly one region section', () => {
    const final = { sectionId: 'final', pageWidth: 12240 };
    const document = documentWith({
      // the parser's `sections` always ends with the final section
      sections: [{ properties: final }],
      finalSectionProperties: final,
    });
    const request = buildResidentRegionLayoutRequest(document, 24, {});
    expect(request.regions.sections).toEqual([{ sectionId: 'final', properties: final }]);
  });

  test('intermediate sections precede the final section, without duplication', () => {
    const intermediate = { sectionId: 'one', pageWidth: 10000 };
    const final = { sectionId: 'final', pageWidth: 12240 };
    const document = documentWith({
      sections: [{ properties: intermediate }, { properties: final }],
      finalSectionProperties: final,
    });
    const request = buildResidentRegionLayoutRequest(document, 24, {});
    expect(request.regions.sections.map((section) => section.sectionId)).toEqual([
      'one',
      'final',
    ]);
  });

  test('the final entry always reflects the editor-maintained final properties', () => {
    const stale = { sectionId: 'final', pageWidth: 12240 };
    const edited = { sectionId: 'final', pageWidth: 11906 };
    const document = documentWith({
      sections: [{ properties: stale }],
      finalSectionProperties: edited,
    });
    const request = buildResidentRegionLayoutRequest(document, 24, {});
    expect(request.regions.sections).toEqual([{ sectionId: 'final', properties: edited }]);
  });

  test('a body without parsed sections still yields the final section', () => {
    const final = { sectionId: 'final' };
    const document = documentWith({ finalSectionProperties: final });
    const request = buildResidentRegionLayoutRequest(document, 24, {});
    expect(request.regions.sections).toEqual([{ sectionId: 'final', properties: final }]);
  });
});
