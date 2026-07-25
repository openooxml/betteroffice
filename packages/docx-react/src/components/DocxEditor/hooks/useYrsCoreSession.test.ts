import { describe, expect, test } from 'bun:test';
import type { Document } from '@betteroffice/docx/types/document';
import type { YrsDocxHost, YrsSession } from '@betteroffice/docx/yrs';
import {
  mergeDocxHostMetadata,
  seedYrsSession,
  warmCompatibilityBase,
} from './useYrsCoreSession';

function fakeSeedSession(): {
  session: Pick<YrsSession, 'openDocx' | 'loadState'>;
  host: YrsDocxHost;
  opened: { bytes: Uint8Array; seedStories: boolean }[];
  loaded: Uint8Array[];
} {
  const host = { name: 'host' } as unknown as YrsDocxHost;
  const opened: { bytes: Uint8Array; seedStories: boolean }[] = [];
  const loaded: Uint8Array[] = [];
  return {
    session: {
      openDocx: (bytes, seedStories) => {
        opened.push({ bytes, seedStories });
        return host;
      },
      loadState: (update) => {
        loaded.push(update);
      },
    },
    host,
    opened,
    loaded,
  };
}

describe('seedYrsSession', () => {
  test('hydrates a pre-parsed document host from the shared initial update', () => {
    const { session, loaded } = fakeSeedSession();
    const initialUpdate = Uint8Array.of(7, 8, 9);
    const seeded: Document[] = [];

    const host = seedYrsSession(session, (document) => seeded.push(document), {
      bytes: null,
      document: { name: 'parsed' } as unknown as Document,
      initialUpdate,
    });

    expect(seeded).toEqual([]);
    expect(loaded).toEqual([initialUpdate]);
    expect(host).toBeNull();
  });

  test('seeds from the pre-parsed document when no shared state exists', () => {
    const { session, loaded } = fakeSeedSession();
    const document = { name: 'parsed' } as unknown as Document;
    const seeded: Document[] = [];

    const host = seedYrsSession(session, (next) => seeded.push(next), {
      bytes: null,
      document,
      initialUpdate: undefined,
    });

    expect(seeded).toEqual([document]);
    expect(loaded).toEqual([]);
    expect(host).toBeNull();
  });

  test('opens bytes for metadata without seeding stories when shared state exists', () => {
    const { session, host: expectedHost, opened, loaded } = fakeSeedSession();
    const bytes = Uint8Array.of(1, 2);
    const initialUpdate = Uint8Array.of(3, 4);

    const host = seedYrsSession(session, () => expect.unreachable(), {
      bytes,
      document: null,
      initialUpdate,
    });

    expect(opened).toEqual([{ bytes, seedStories: false }]);
    expect(loaded).toEqual([initialUpdate]);
    expect(host).toBe(expectedHost);
  });

  test('seeds stories from bytes when no shared state exists', () => {
    const { session, host: expectedHost, opened, loaded } = fakeSeedSession();
    const bytes = Uint8Array.of(1, 2);

    const host = seedYrsSession(session, () => expect.unreachable(), {
      bytes,
      document: null,
      initialUpdate: undefined,
    });

    expect(opened).toEqual([{ bytes, seedStories: true }]);
    expect(loaded).toEqual([]);
    expect(host).toBe(expectedHost);
  });
});

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

describe('warmCompatibilityBase', () => {
  test('materializes the projection base once', () => {
    let materializations = 0;
    const materialized = { name: 'materialized' } as unknown as Document;
    const session = {
      materializeDocx: () => {
        materializations += 1;
        return materialized;
      },
    };
    const compatibilityBase: { current: Document | null } = { current: null };

    warmCompatibilityBase(session, compatibilityBase);
    warmCompatibilityBase(session, compatibilityBase);

    expect(compatibilityBase.current).toBe(materialized);
    expect(materializations).toBe(1);
  });

  test('leaves an already-projected base untouched', () => {
    const projected = { name: 'projected' } as unknown as Document;
    const compatibilityBase: { current: Document | null } = { current: projected };

    warmCompatibilityBase({ materializeDocx: () => expect.unreachable() }, compatibilityBase);

    expect(compatibilityBase.current).toBe(projected);
  });
});
