import { describe, expect, test } from 'bun:test';
import type { Document } from '@betteroffice/docx/types/document';
import { HistoryManager } from '../../hooks/useHistory';
import {
  commitLegacyDocumentChange,
  commitYrsDocumentChange,
  type DocumentChangeSinks,
  type ProjectionScheduler,
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

function manualScheduler(): { schedule: ProjectionScheduler; flush: () => void } {
  let pending: (() => void) | null = null;
  return {
    schedule: (project) => {
      pending = project;
    },
    flush: () => {
      const project = pending;
      pending = null;
      project?.();
    },
  };
}

describe('document change identity', () => {
  test('stores and notifies the same projection for a legacy page-setup change', () => {
    const legacyDocument = { name: 'legacy' } as unknown as Document;
    const projection = { name: 'projected' } as unknown as Document;
    const history = new HistoryManager<Document | null>(null);
    const { sinks, notified } = historySinks(history);
    const scheduler = manualScheduler();

    const committed = commitLegacyDocumentChange(
      legacyDocument,
      (baseDocument) => {
        expect(baseDocument).toBe(legacyDocument);
        return projection;
      },
      sinks,
      scheduler.schedule
    );

    expect(committed).toBe(legacyDocument);
    expect(history.state).toBe(legacyDocument);
    scheduler.flush();

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

describe('legacy change projection cost', () => {
  test('never projects while nothing listens for changes', () => {
    const history = new HistoryManager<Document | null>(null);
    const scheduler = manualScheduler();
    let projections = 0;
    const project = () => {
      projections += 1;
      return { name: 'projected' } as unknown as Document;
    };

    const drag = [1, 2, 3, 4, 5].map(
      (margin) => ({ name: `margin-${margin}` }) as unknown as Document
    );
    for (const document of drag) {
      commitLegacyDocumentChange(
        document,
        project,
        { push: (next) => history.push(next) },
        scheduler.schedule
      );
    }
    scheduler.flush();

    expect(projections).toBe(0);
    expect(history.state).toBe(drag[drag.length - 1]);
  });

  test('coalesces a drag into one projection of the latest document', () => {
    const history = new HistoryManager<Document | null>(null);
    const { sinks, notified } = historySinks(history);
    const scheduler = manualScheduler();
    const projectedBases: (Document | null | undefined)[] = [];
    const project = (baseDocument?: Document | null) => {
      projectedBases.push(baseDocument);
      return { name: 'projected' } as unknown as Document;
    };

    const drag = [1, 2, 3, 4, 5].map(
      (margin) => ({ name: `margin-${margin}` }) as unknown as Document
    );
    for (const document of drag) {
      commitLegacyDocumentChange(document, project, sinks, scheduler.schedule);
      expect(projectedBases).toHaveLength(0);
      expect(history.state).toBe(document);
    }
    scheduler.flush();

    expect(projectedBases).toEqual([drag[drag.length - 1]]);
    expect(notified()).toBe(history.state);
  });

  test('projects the first commit so edits made before it survive', () => {
    const history = new HistoryManager<Document | null>(null);
    const { sinks, notified } = historySinks(history);
    const scheduler = manualScheduler();
    const legacyDocument = { name: 'legacy', text: '' } as unknown as Document;
    const project = (baseDocument?: Document | null) =>
      ({
        ...(baseDocument as object),
        text: 'typed before the first commit',
      }) as unknown as Document;

    commitLegacyDocumentChange(legacyDocument, project, sinks, scheduler.schedule);
    scheduler.flush();

    expect((notified() as unknown as { text: string }).text).toBe('typed before the first commit');
    expect(notified()).toBe(history.state);
  });
});
