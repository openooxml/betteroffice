import { describe, expect, test } from 'bun:test';
import type { Document } from '@betteroffice/docx/types/document';
import { mergeDocxHostMetadata } from './useYrsCoreSession';

describe('mergeDocxHostMetadata', () => {
  test('preserves recursive content while applying live host metadata', () => {
    const bodyContent = [
      { type: 'paragraph', content: [] },
    ] as unknown as Document['package']['document']['content'];
    const sectionContent = [
      { type: 'paragraph', content: [] },
    ] as unknown as Document['package']['document']['content'];
    const headerContent = [
      { type: 'paragraph', content: [] },
    ] as unknown as Document['package']['document']['content'];
    const noteContent = [
      { type: 'paragraph', content: [] },
    ] as unknown as Document['package']['document']['content'];
    const media = new Map([
      ['word/media/image1.png', { path: 'word/media/image1.png' }],
    ]) as unknown as NonNullable<Document['package']['media']>;
    const savedBuffer = Uint8Array.of(1).buffer;
    const sourceBuffer = Uint8Array.of(2).buffer;
    const full = {
      originalBuffer: savedBuffer,
      package: {
        document: {
          content: bodyContent,
          sections: [
            {
              id: 'section-1',
              properties: { marginTop: 100 },
              content: sectionContent,
            },
          ],
          finalSectionProperties: { marginTop: 100 },
        },
        headers: new Map([
          [
            'rId1',
            {
              type: 'header',
              hdrFtrType: 'default',
              content: headerContent,
            },
          ],
        ]),
        footnotes: [
          {
            type: 'footnote',
            id: 1,
            noteType: 'normal',
            content: noteContent,
          },
        ],
        media,
      },
    } as unknown as Document;
    const relationships = new Map();
    const host = {
      originalBuffer: sourceBuffer,
      package: {
        document: {
          content: [],
          sections: [
            {
              id: 'section-1',
              properties: { marginTop: 720 },
              content: [],
            },
          ],
          finalSectionProperties: { marginTop: 720 },
        },
        headers: new Map([
          [
            'rId1',
            {
              type: 'header',
              hdrFtrType: 'first',
              content: [],
            },
          ],
        ]),
        footnotes: [
          {
            type: 'footnote',
            id: 1,
            noteType: 'normal',
            content: [],
          },
        ],
        relationships,
      },
    } as unknown as Document;

    const merged = mergeDocxHostMetadata(full, host);

    expect(merged.package.document.content).toBe(bodyContent);
    expect(merged.package.document.sections?.[0].content).toBe(sectionContent);
    expect(merged.package.document.sections?.[0].properties.marginTop).toBe(720);
    expect(merged.package.headers?.get('rId1')?.content).toBe(headerContent);
    expect(merged.package.headers?.get('rId1')?.hdrFtrType).toBe('first');
    expect(merged.package.footnotes?.[0].content).toBe(noteContent);
    expect(merged.package.relationships).toBe(relationships);
    expect(merged.package.media).toBe(media);
    expect(merged.originalBuffer).toBe(savedBuffer);
  });
});
