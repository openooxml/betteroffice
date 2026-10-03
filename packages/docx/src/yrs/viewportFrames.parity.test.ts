import { beforeAll, expect, test } from 'bun:test';
import { readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';

import { buildResidentRegionLayoutRequest } from '../editor/computeLayout';
import type { DisplayPage } from '../layout/render/displayList';
import {
  applyFrameDelta,
  applyFrameDeltaOwned,
  decodeFrameDelta,
  type DecodedFrameDelta,
  type RetainedFrame,
} from '../layout/render/frameDelta';
import {
  encodeDisplayListFrameExtras,
  type DisplayListBuildInputs,
} from '../layout/render/rustDisplayList';
import { preloadEditWasm } from '../wasm/edit';
import { syntheticDocx, type Flavour } from './__fixtures__/previewChain';
import { createYrsSession, decodeDocxHostJson } from './index';
import { residentCaretSnapshotForFrame } from './residentCaret';
import { createResidentEngineSession, type ResidentEngineSession } from './residentEngineSession';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const CORPORA = [
  'crates/docx-edit/tests/fixtures',
  'crates/betteroffice-docx/tests/corpus/fixtures',
];
const FONT = new Uint8Array(
  readFileSync(
    resolve(import.meta.dir, '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf')
  )
);
const SEED = 0x72616e67;
const STEPS = 12;
const SMALL_FIXTURE_STEPS = 6;
const RECOVERY_STEP = 1;
const BASELINE_EXCLUSIONS: Array<[string, string]> = [];
const FLAVOURS: Flavour[] = [
  'plain',
  'keepnext',
  'widows',
  'footnotes',
  'endnotes',
  'floats',
  'early-sect',
];
const NO_WINDOW_FLAVOURS: Flavour[] = ['plain', 'footnotes'];

interface Fixture {
  name: string;
  bytes: Uint8Array;
  seed: number;
  corpus: boolean;
  noWindow?: boolean;
}

interface Hosts {
  owned: RetainedFrame | null;
  copying: RetainedFrame | null;
}

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

function mulberry32(seed: number): () => number {
  return () => {
    seed = (seed + 0x6d2b79f5) >>> 0;
    let value = Math.imul(seed ^ (seed >>> 15), seed | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 2 ** 32;
  };
}

function* fixtures(): Generator<Fixture> {
  for (const [index, flavour] of FLAVOURS.entries()) {
    const seed = (SEED + index) >>> 0;
    const fixture = {
      name: `synthetic/${flavour}`,
      bytes: syntheticDocx(flavour, 48, seed),
      seed,
      corpus: false,
    };
    yield fixture;
    if (NO_WINDOW_FLAVOURS.includes(flavour)) {
      yield { ...fixture, name: `${fixture.name} (no window)`, noWindow: true };
    }
  }
  yield {
    name: 'synthetic/page-restarts',
    bytes: syntheticDocx('plain', 48, SEED + 7, {
      pageNumberRestarts: [1, 7],
    }),
    seed: SEED + 7,
    corpus: false,
  };
  let index = 0;
  for (const corpus of CORPORA) {
    const directory = resolve(import.meta.dir, '../../../../', corpus);
    const names = readdirSync(directory, { recursive: true, encoding: 'utf8' })
      .filter((name) => name.endsWith('.docx'))
      .sort();
    for (const name of names) {
      yield {
        name: `${corpus}/${name}`,
        bytes: new Uint8Array(readFileSync(resolve(directory, name))),
        seed: (SEED + 100 + index++) >>> 0,
        corpus: true,
      };
    }
  }
}

function extras(session: ResidentEngineSession, fontChains: Record<string, number[]>): string {
  const retained = session.retainedHeadersFootersJson();
  return encodeDisplayListFrameExtras({
    fontChains,
    ...(retained === undefined ? {} : { headersFooters: JSON.parse(retained) }),
  } as DisplayListBuildInputs);
}

function present(
  session: ResidentEngineSession,
  hosts: Hosts,
  bytes: Uint8Array,
  context: string
): DecodedFrameDelta {
  const delta = decodeFrameDelta(bytes);
  hosts.owned = applyFrameDeltaOwned(hosts.owned, delta);
  hosts.copying = applyFrameDelta(hosts.copying, decodeFrameDelta(bytes));
  const caret = session.residentCaretSnapshot();
  expect(
    residentCaretSnapshotForFrame(caret, hosts.owned),
    `${context} page=-1 key=caret`
  ).not.toBeNull();
  return delta;
}

function firstDifference(actual: unknown, expected: unknown, path = 'page'): string | null {
  if (Object.is(actual, expected)) return null;
  if (
    actual === null ||
    expected === null ||
    typeof actual !== 'object' ||
    typeof expected !== 'object'
  ) {
    return path;
  }
  if (Array.isArray(actual) !== Array.isArray(expected)) return path;
  if (Array.isArray(actual) && Array.isArray(expected) && actual.length !== expected.length) {
    return `${path}.length`;
  }
  const left = actual as Record<string, unknown>;
  const right = expected as Record<string, unknown>;
  for (const key of new Set([...Object.keys(left), ...Object.keys(right)])) {
    const next = Array.isArray(actual) ? `${path}[${key}]` : `${path}.${key}`;
    if (Object.hasOwn(left, key) !== Object.hasOwn(right, key)) return next;
    const difference = firstDifference(left[key], right[key], next);
    if (difference) return difference;
  }
  return null;
}

function expectPages(
  hosts: Hosts,
  oracle: RetainedFrame,
  context: string,
  window: [number, number] | undefined
): void {
  for (const [chain, frame] of Object.entries(hosts)) {
    expect(frame, `${context} chain=${chain} page=-1 key=frame`).not.toBeNull();
    const pages = frame!.displayList.pages;
    expect(pages.length, `${context} chain=${chain} page=-1 key=pages.length`).toBe(
      oracle.displayList.pages.length
    );
    for (const [index, expected] of oracle.displayList.pages.entries()) {
      if (window === undefined || (index >= window[0] && index < window[1])) {
        expect(
          pages[index]?.unbuilt,
          `${context} chain=${chain} page=${index} key=unbuilt`
        ).not.toBe(true);
      }
      try {
        expect(pages[index], `${context} chain=${chain} page=${index}`).toEqual(expected);
      } catch (error) {
        throw new Error(
          `${context} chain=${chain} page=${index} key=${firstDifference(
            pages[index],
            expected
          )}\n${String(error)}`
        );
      }
    }
  }
}

async function coldFrame(
  session: ResidentEngineSession,
  input: string,
  fontChains: Record<string, number[]>,
  mediaSources: string,
  noteSeparators: Uint8Array,
  pages: DisplayPage[]
): Promise<RetainedFrame> {
  const cold = await createResidentEngineSession();
  try {
    cold.loadState(session.encodeState());
    cold.loadMediaSources(mediaSources);
    cold.loadNoteSeparators(noteSeparators);
    cold.registerFont(FONT);
    cold.layoutDocumentWithRegionsRetained(input);
    let oracle = applyFrameDelta(
      null,
      decodeFrameDelta(cold.buildDisplayListFrame(extras(cold, fontChains), 0))
    );
    const unbuilt = pages.filter((page) => page.unbuilt === true).map((page) => page.pageIndex);
    if (unbuilt.length > 0) {
      const released = cold.releaseDisplayPagesFrame(unbuilt, oracle.frameEpoch);
      if (released === null) throw new Error('Oracle refused to release display pages');
      oracle = applyFrameDelta(oracle, decodeFrameDelta(released));
    }
    return oracle;
  } finally {
    cold.destroy();
  }
}

test('viewport frames match a cold rebuild after every resident edit and scroll', async () => {
  const started = performance.now();
  let documents = 0;
  let corpusDocuments = 0;
  let attempted = 0;
  let totalSteps = 0;
  let smallDocuments = 0;
  let positionFrames = 0;
  let rangeFrames = 0;
  const refused: string[] = [];
  const excluded: string[] = [];
  const failures: string[] = [];
  for (const fixture of fixtures()) {
    const { name, bytes, seed, corpus, noWindow = false } = fixture;
    attempted += 1;
    const exclusion = BASELINE_EXCLUSIONS.find(([excluded]) => excluded === name);
    const session = await createResidentEngineSession();
    let context = `document=${name} seed=${seed} step=-1 op=initial`;
    try {
      let hostJson: string;
      try {
        session.setDirectBatches(true);
        hostJson = session.openDocx(bytes);
      } catch (error) {
        if (corpus && /\b(refused|unsupported)\b/i.test(String(error))) {
          refused.push(`${name}: ${String(error)}`);
          continue;
        }
        throw error;
      }
      const document = decodeDocxHostJson(hostJson, bytes).document;
      const themeColors = Object.fromEntries(
        Object.entries(document.package.theme?.colorScheme ?? {}).filter(
          (entry): entry is [string, string] => typeof entry[1] === 'string'
        )
      );
      const request = buildResidentRegionLayoutRequest(document, 24, {
        themeColors,
        defaultTabStopTwips: document.package.settings?.defaultTabStop ?? null,
        numericIds: {},
        showHiddenText: false,
        mediaTokens: true,
      });
      const requirements = JSON.parse(
        session.layoutFontRequirementsJson(JSON.stringify(request))
      ) as Array<{ key: string }>;
      const font = session.registerFont(FONT);
      const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [font]]));
      const compat = document.package.settings?.compatibilityFlags;
      request.measurement = {
        fontChains,
        defaults: { fontSize: 11, fontFamily: 'Calibri' },
        compat: {
          noLeading: compat?.noLeading ?? false,
          doNotExpandShiftReturn: compat?.doNotExpandShiftReturn ?? false,
        },
        authoritativeShaping: true,
      };
      const input = JSON.stringify(request);
      let mediaSources = '';
      let noteSeparators: Uint8Array = new Uint8Array(0);
      const main = await createYrsSession();
      try {
        main.openDocx(bytes, true);
        main.adoptResidentWorkerLayout!(input);
        const snapshot = main.residentWorkerSnapshot()!;
        mediaSources = snapshot.mediaSources ?? '';
        noteSeparators = snapshot.noteSeparators ?? new Uint8Array(0);
      } finally {
        main.destroy();
      }
      let window: [number, number] = [0, 3];
      const hosts: Hosts = { owned: null, copying: null };
      const retainBuiltPages = seed % 3 === 0;
      const setWindow = () => {
        if (!noWindow) session.setDisplayWindow(...window);
        session.setDisplayRetainBuiltPages(retainBuiltPages);
        session.setWindowedIncrementalBuilds(!noWindow);
      };
      setWindow();
      session.layoutDocumentWithRegionsRetained(input);
      present(
        session,
        hosts,
        session.buildDisplayListFrame(extras(session, fontChains), 0),
        context
      );
      const pageCount = hosts.owned!.displayList.pages.length;
      const steps = corpus && pageCount < 3 ? SMALL_FIXTURE_STEPS : STEPS;
      if (!corpus) {
        expect(pageCount, `${context} page=-1 key=pages.length`).toBeGreaterThanOrEqual(15);
        expect(pageCount, `${context} page=-1 key=pages.length`).toBeLessThanOrEqual(40);
      }
      if (name === 'synthetic/page-restarts') {
        const pages = hosts.owned!.displayList.pages;
        expect(
          new Set(pages.map((page) => page.sectionIndex)).size,
          `${context} page=-1 key=sections`
        ).toBe(2);
        expect(pages[0]?.sectionPageNumber, `${context} page=0 key=sectionPageNumber`).toBe(1);
        const restart = pages.find((page) => page.sectionIndex === 1)!;
        expect(
          restart.sectionPageNumber,
          `${context} page=${restart.pageIndex} key=sectionPageNumber`
        ).toBe(7);
        expect(
          pages[0]?.footer?.primitives.some(
            (primitive) => 'field' in primitive && primitive.field?.category === 'PAGE'
          ),
          `${context} page=0 key=footer.PAGE`
        ).toBe(true);
      }
      const baseline = await coldFrame(
        session,
        input,
        fontChains,
        mediaSources,
        noteSeparators,
        hosts.owned!.displayList.pages
      );
      try {
        expectPages(hosts, baseline, context, noWindow ? undefined : window);
      } catch (error) {
        if (!exclusion) throw error;
        excluded.push(`${name}: ${exclusion[1]}`);
        continue;
      }
      const expectColdPages = async () => {
        const oracle = await coldFrame(
          session,
          input,
          fontChains,
          mediaSources,
          noteSeparators,
          hosts.owned!.displayList.pages
        );
        expectPages(
          hosts,
          oracle,
          context,
          noWindow ? undefined : window
        );
        return oracle;
      };
      const random = mulberry32(seed);
      const pick = (limit: number) => Math.floor(random() * limit);
      for (let step = 0; step < steps; step++) {
        if (step === 4 || step === 7 || step === 11) {
          const count = hosts.owned!.displayList.pages.length;
          const start = pick(count);
          window = [start, Math.min(count, start + 1 + pick(3))];
          context = `document=${name} seed=${seed} step=${step} op=scroll(${window.join('..')})`;
          setWindow();
          const indices = Array.from(
            { length: window[1] - window[0] },
            (_, index) => window[0] + index
          );
          present(
            session,
            hosts,
            session.buildDisplayPagesFrame(indices, hosts.owned!.frameEpoch),
            context
          );
          await expectColdPages();
          const release = hosts
            .owned!.displayList.pages.filter(
              (page) =>
                !noWindow &&
                !page.unbuilt &&
                (page.pageIndex < window[0] || page.pageIndex >= window[1])
            )
            .map((page) => page.pageIndex);
          if (release.length > 0) {
            const frame = session.releaseDisplayPagesFrame(release, hosts.owned!.frameEpoch);
            expect(frame, `${context} page=-1 key=release`).not.toBeNull();
            present(session, hosts, frame!, context);
            await expectColdPages();
          }
        } else {
          const paragraphs = session.geometryReader.positionOutline!('body')!.body!.paragraphs;
          const mergeTargets = paragraphs.filter(
            (paragraph, index) => index > 0 && paragraph.leading === 0
          );
          const merge = step === 8 && mergeTargets.length > 0;
          const targets = merge ? mergeTargets : paragraphs;
          const paragraph = targets[pick(targets.length)]!;
          const localOffset = merge ? 0 : pick(paragraph.length + 1);
          const offset = paragraph.leading + localOffset;
          const head = { story: 'body', paraId: paragraph.paraId, offset };
          const enter = step === 3;
          const deleting = merge || (!enter && localOffset > 0 && random() < 0.4);
          const count = merge ? 1 + pick(2) : Math.min(localOffset, 1 + pick(2));
          const text = enter
            ? '\n'
            : Array.from({ length: 1 + pick(3) }, () => 'abc xyz'[pick(7)]).join('');
          const op = deleting
            ? `${merge ? 'merge' : 'backspace'}(${count})`
            : enter
            ? 'enter'
            : `insert(${JSON.stringify(text)})`;
          context = `document=${name} seed=${seed} step=${step} op=${op}@${head.paraId}:${offset}`;
          setWindow();
          session.setSelection(head, head);
          if (enter) {
            const main = await createYrsSession();
            try {
              main.loadState(session.encodeState());
              const split = main.splitParagraph(head);
              session.applyUpdate(main.encodeStateAsUpdate(session.encodeStateVector()));
              session.setSelection({ story: 'body', paraId: split.secondParaId, offset: 0 });
              session.layoutDocumentWithRegionsRetained(input);
              present(
                session,
                hosts,
                session.buildDisplayListFrame(extras(session, fontChains), hosts.owned!.frameEpoch),
                context
              );
            } finally {
              main.destroy();
            }
          } else {
            const delta = present(
              session,
              hosts,
              deleting
                ? session.applyDelete('backward', hosts.owned!.frameEpoch, count)
                : session.applyInput(text, hosts.owned!.frameEpoch),
              context
            );
            expect(delta.full, `${context} page=-1 key=full`).toBe(false);
            if (
              delta.operations.some(
                ({ kind }) => kind.startsWith('shift-') || kind === 'patch-positions'
              )
            ) {
              positionFrames += 1;
            }
            if (delta.operations.some(({ kind }) => kind === 'shift-range')) rangeFrames += 1;
          }
          const oracle = await expectColdPages();
          if (step === RECOVERY_STEP) {
            context = `${context} frame=recovery`;
            const delta = present(session, hosts, session.buildDisplayPagesFrame([], 0), context);
            expect(delta.full, `${context} page=-1 key=full`).toBe(true);
            expectPages(hosts, oracle, context, noWindow ? undefined : window);
          }
        }
        totalSteps += 1;
      }
      documents += 1;
      if (corpus) {
        corpusDocuments += 1;
        if (pageCount < 3) smallDocuments += 1;
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      failures.push(
        message.startsWith(context)
          ? message.split('\n')[0]!
          : `${context} page=-1 key=session: ${message.split('\n')[0]}`
      );
    } finally {
      session.destroy();
    }
  }
  console.info(
    `viewport frames: ${documents}/${attempted} documents (${corpusDocuments} corpus, ${smallDocuments} small at ${SMALL_FIXTURE_STEPS} steps, others at ${STEPS}), ${totalSteps} steps, ${positionFrames} position frames (${rangeFrames} with range shifts), ${(
      (performance.now() - started) / 1000
    ).toFixed(2)} s; refused: ${refused.join('; ') || 'none'}; exclusions: ${excluded.join('; ') || 'none'}; failures: ${failures.join('; ') || 'none'}`
  );
  expect(failures, 'viewport frame parity failures').toEqual([]);
  expect(
    documents + refused.length + excluded.length,
    `document=selection seed=${SEED} step=-1 op=coverage page=-1 key=documents`
  ).toBe(attempted);
  expect(
    documents,
    `document=selection seed=${SEED} step=-1 op=coverage page=-1 key=documents`
  ).toBeGreaterThanOrEqual(
    FLAVOURS.length +
      1 +
      NO_WINDOW_FLAVOURS.length -
      excluded.filter((name) => name.startsWith('synthetic/')).length
  );
  expect(
    positionFrames,
    `document=selection seed=${SEED} step=-1 op=coverage page=-1 key=positionFrames`
  ).toBeGreaterThan(0);
  expect(
    rangeFrames,
    `document=selection seed=${SEED} step=-1 op=coverage page=-1 key=rangeFrames`
  ).toBeGreaterThan(0);
}, 120_000);
