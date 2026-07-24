import { describe, expect, test } from 'bun:test';
import type { Document } from '@betteroffice/docx/types/document';
import { HistoryManager } from '../../hooks/useHistory';
import {
  commitLegacyDocumentChange,
  commitYrsDocumentChange,
  type DocumentChangeSinks,
} from './documentChangeCommit';

function historySinks(history: HistoryManager<Document | null>): {
  sinks: DocumentChangeSinks;
  notified: () => Document | null;
} {
  let notified: Document | null = null;
  return {
    sinks: {
      push: (document) => history.push(document),
      notify: (document) => {
        notified = document;
      },
    },
    notified: () => notified,
  };
}

describe('document change identity', () => {
  test('stores and notifies the same projection for a legacy page-setup change', () => {
    const legacyDocument = { name: 'legacy' } as unknown as Document;
    const projection = { name: 'projected' } as unknown as Document;
    const history = new HistoryManager<Document | null>(null);
    const { sinks, notified } = historySinks(history);

    const committed = commitLegacyDocumentChange(
      legacyDocument,
      (baseDocument) => {
        expect(baseDocument).toBe(legacyDocument);
        return projection;
      },
      sinks
    );

    expect(committed).toBe(projection);
    expect(notified()).toBe(history.state);
    expect(history.state).toBe(projection);
  });

  test('stores and notifies the same projection for a yrs-origin change', () => {
    const projection = { name: 'projected' } as unknown as Document;
    const history = new HistoryManager<Document | null>(null);
    const { sinks, notified } = historySinks(history);

    const committed = commitYrsDocumentChange(() => projection, sinks);

    expect(committed).toBe(projection);
    expect(notified()).toBe(history.state);
    expect(history.state).toBe(projection);
  });
});
