import { afterAll, beforeAll, describe, expect, it } from 'bun:test';

import type { Document } from '../types/document';
import { dirtyProjectionStory } from './dirtyProjectionStories';
import type { YrsSession } from './index';
import { projectionAllowlist } from './projectionParity.allowlist';
import {
  buildFixtureDocx,
  canonicalStories,
  check,
  checkComposition,
  checkStory,
  compareSnapshots,
  createEditSession,
  edit,
  nativeProjection,
  nativeStoryProjection,
  preloadProjectionWasm,
  projectionExport,
  projectionRootOrder,
  random,
  seedParityFixtures,
  type AllowanceUsage,
  type NativeProjection,
  type ProjectionEdit,
} from './projectionParity.support';
import { wrapSession } from './yrsSessionFacade';
import { yrsToDocument } from './yrsToDocument';

const CLIENT = 48001;
const expectedTypeScriptDisagreement = {
  fixture: 'page-break-beside-field',
  batch: 4,
  reason: 'Full and incremental TypeScript group text typed at a bold run boundary differently.',
  differences: [
    'story body, paragraph 0, blocks[0].content[0].content[0].text: TS "InQ9ZrQ9Z"; native "InQ9Zr"',
    'story body, paragraph 0, blocks[0].content[1].formatting: TS {"bold":true}; native <missing>',
    'story body, paragraph 0, blocks[0].content[1].content[0].text: TS "o"; native "Q9Z"',
    'story body, paragraph 0, blocks[0].content[2]: TS <missing>; native {"content":[{"text":"o","type":"text"}],"formatting":{"bold":true},"type":"run"}',
  ],
};

function checkDisagreementFixture(fixtures: readonly string[]): void {
  expect(fixtures.length).toBeGreaterThan(0);
  expect(fixtures).toContain(expectedTypeScriptDisagreement.fixture);
}

it('rejects discovery that omits the expected TypeScript disagreement fixture', () => {
  expect(() => checkDisagreementFixture(['another-fixture'])).toThrow();
});

function editParagraphProperties(session: YrsSession, recordEdit: (outcome: ProjectionEdit) => void): void {
  const paragraphs = session.paragraphs('body');
  const paragraph = paragraphs[2] ?? paragraphs[0];
  const outcome = { kind: 'paragraph properties', story: 'body', paraId: paragraph?.paraId ?? '' };
  if (!paragraph?.paraId) {
    recordEdit({ ...outcome, refusal: 'no body paragraph' });
    return;
  }
  const position = { paraId: paragraph.paraId, offset: 0 };
  const version = session.version();
  try {
    session.setParagraphAttrs({ story: 'body', start: position, end: position }, {
      lineSpacing: 360,
      lineSpacingRule: 'auto',
      indentLeft: 1080,
    });
  } catch (error) {
    recordEdit({ ...outcome, refusal: String(error) });
    return;
  }
  recordEdit(version === session.version() ? { ...outcome, refusal: 'no state change' } : outcome);
}

describe('Native DOCX projection parity', () => {
  const fixtures = seedParityFixtures();
  const used: AllowanceUsage = new Map();
  let checkedDisagreement = false;

  beforeAll(async () => {
    checkDisagreementFixture(fixtures);
    await preloadProjectionWasm();
    const raw = createEditSession(CLIENT);
    try {
      projectionExport(raw, 'project_document_json');
      projectionExport(raw, 'project_story_json');
    } finally {
      raw.free();
    }
  }, 30_000);

  for (const name of fixtures) {
    it(`matches the TypeScript projection of ${name} through deterministic edits`, () => {
      const raw = createEditSession(CLIENT);
      const session = wrapSession(raw, CLIENT);
      try {
        const bytes = buildFixtureDocx(name);
        session.seedFromDocx(bytes);
        const base = session.materializeDocx();
        if (!base) throw new Error('DOCX projection requires the complete source document');
        const buffer = base.originalBuffer!;
        const source = canonicalStories(session, base);
        const next = random(name);
        const failures: string[] = [];
        const checkState = (label: string, differences: string[]) => {
          if (differences.length) failures.push(`${label}\n${differences.join('\n')}`);
        };
        const checkProjections = (label: string, ts: Document, native: NativeProjection): void => {
          checkState(`${label}, document`, check(session, ts, native, source, undefined, used));
          const expected = canonicalStories(session, ts);
          const roots = projectionRootOrder([...expected.keys()].filter((id) => dirtyProjectionStory(id) === id));
          const stories = roots.map((id) => {
            const story = nativeStoryProjection(raw, id);
            checkState(`${label}, ${id}`, checkStory(expected, story, source, used));
            return story;
          });
          checkState(`${label}, composed stories versus document`, checkComposition(session, stories, native));
        };

        let ts = yrsToDocument(session, base);
        let native = nativeProjection(raw, buffer);
        checkProjections('seeded', ts, native);

        const edits: ProjectionEdit[] = [];
        const knownDisagreement = name === expectedTypeScriptDisagreement.fixture ? expectedTypeScriptDisagreement : undefined;
        let sawDisagreement = false;
        for (let trial = 0; trial < 5; trial++) {
          const outcomes: ProjectionEdit[] = [];
          const recordEdit = (outcome: ProjectionEdit) => outcomes.push(outcome);
          const before = session.storiesChangedSince(Number.MAX_SAFE_INTEGER).revision;
          if (trial === 4) editParagraphProperties(session, recordEdit);
          else edit(session, next, recordEdit);
          edits.push(...outcomes);
          const changed = session.storiesChangedSince(before).stories;
          const roots = new Set(changed.map(dirtyProjectionStory));
          const untouched = new Set([...canonicalStories(session, ts).keys()]
            .map(dirtyProjectionStory).filter((story) => !roots.has(story)));
          const label = `edit batch ${trial + 1} (${outcomes.map(({ kind, story, paraId, refusal }) =>
            `${kind}@${story}/${paraId} ${refusal ? `refused: ${refusal}` : 'applied'}`
          ).join(', ') || 'no eligible edit targets'})`;
          if (roots.size > 0) {
            const incrementalTs = yrsToDocument(session, ts, { storyIds: roots });
            const incrementalNative = nativeProjection(raw, buffer, roots);
            checkState(`${label}, untouched TypeScript roots`, compareSnapshots(session, ts, incrementalTs, untouched));
            checkState(`${label}, untouched native roots`, compareSnapshots(session, native.document, incrementalNative.document, untouched));

            ts = yrsToDocument(session, base);
            native = nativeProjection(raw, buffer);
            checkState(`${label}, selected roots`, check(session, ts, incrementalNative, source, roots, used));
            const disagreement = compareSnapshots(session, ts, incrementalTs, undefined);
            if (trial + 1 === knownDisagreement?.batch) {
              expect(disagreement, knownDisagreement.reason).toEqual(knownDisagreement.differences);
              sawDisagreement = true;
              checkedDisagreement = true;
            } else {
              checkState(`${label}, full TypeScript projection`, disagreement);
            }
            checkState(`${label}, full native projection`, compareSnapshots(session, native.document, incrementalNative.document, undefined));
          } else {
            ts = yrsToDocument(session, base);
            native = nativeProjection(raw, buffer);
          }
          checkProjections(label, ts, native);
        }
        expect(edits.some((outcome) => !outcome.refusal), JSON.stringify(edits)).toBe(true);
        expect(failures, failures.join('\n\n')).toEqual([]);
        if (knownDisagreement) expect(sawDisagreement, knownDisagreement.reason).toBe(true);
      } finally {
        session.destroy();
      }
    }, 30_000);
  }

  afterAll(() => {
    const owners = {
      fields: 'P1.1 fields',
      tables: 'P1.2 tables',
      markers: 'P1.3 markers',
      storiesAux: 'P1.4 storiesAux',
    } as const;
    for (const family of Object.keys(owners) as (keyof typeof owners)[]) {
      const entries = projectionAllowlist.filter((entry) => entry.family === family);
      const covered = entries.reduce((count, entry) => count + (used.get(entry) ?? 0), 0);
      console.log(`${owners[family]} covers ${covered} differences and deferrals with ${entries.length} allowlist entries`);
    }
    const unused = projectionAllowlist.filter((entry) => !used.has(entry));
    const details = unused.map(({ family, kind, path, reason }) => `${family}/${kind} ${path}: ${reason}`);
    expect(unused, details.join('\n')).toEqual([]);
    expect(checkedDisagreement, expectedTypeScriptDisagreement.reason).toBe(true);
  });
});
