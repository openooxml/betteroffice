import { afterEach, expect, test } from 'bun:test';
import { isEngineTrap } from '../../../../shared/office-session/trap';
import type { YrsCellLoc, YrsStorySegment } from './index';
import * as selection from './residentCommandTableSelection';
const { resolveLiveCell } = selection;
import type { DocxWorkerCommandEngine } from './residentCommandProtocol';
import { closeCommandEngines, openCommandEngine } from './residentCommandTestSupport';
import { createYrsPositionProjection } from './yrsPositionProjection';

const SEEDS = [11, 23, 37, 41, 59];
const STEPS = 300;

const w15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string) => `<w:tc><w:tcPr/>${content}</w:tc>`;
const table = (rows: string[][]) => `<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="1500"/>'.repeat(rows[0]!.length)}</w:tblGrid>${
  rows.map((row) => `<w:tr>${row.map(cell).join('')}</w:tr>`).join('')}</w:tbl>`;
const grid = (prefix: string, rows: number, columns: number) =>
  table(Array.from({ length: rows }, (_, r) => Array.from({ length: columns }, (_, c) => paragraph(`${prefix}${r}${c}`))));
const control = `<w:sdt><w:sdtPr><w:alias w:val="Box"/></w:sdtPr><w:sdtContent>${paragraph('in control')}${grid('X', 2, 2)}</w:sdtContent></w:sdt>`;
const repeat = `<w:sdt><w:sdtPr><w:id w:val="8"/><w15:repeatingSection ${w15}/></w:sdtPr><w:sdtContent><w:sdt><w:sdtPr><w:id w:val="9"/><w:alias w:val="Line"/><w15:repeatingSectionItem ${w15}/></w:sdtPr><w:sdtContent>${grid('R', 1, 2)}${paragraph('Item')}</w:sdtContent></w:sdt></w:sdtContent></w:sdt>`;
const body = paragraph('Before') +
  table([[paragraph('A'), grid('N', 2, 2) + paragraph(''), paragraph('C')], [paragraph('D'), paragraph('E'), paragraph('F')], [paragraph('G'), paragraph('H'), paragraph('I')]]) +
  paragraph('Middle') + control + repeat + paragraph('After');

afterEach(closeCommandEngines);

function random(seed: number) {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** The engine without its cell shortcuts: resolveLiveCell projects on every call, as before any cache or engine read. */
function withoutCellLocation(engine: DocxWorkerCommandEngine): DocxWorkerCommandEngine {
  return new Proxy(engine, { get: (target, key) => (key === 'cellLocation' || key === 'cellStructureKey' ? undefined : Reflect.get(target, key)) });
}

async function packageBytes(bodyXml: string): Promise<Uint8Array> {
  const { rezipPartsToArrayBuffer, toBytes } = await import('../docx/rezip/parts');
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const units = (segment: YrsStorySegment) => (segment.kind === 'text' ? segment.text.length : 1);

function embeds(engine: DocxWorkerCommandEngine, story: string, kind: string) {
  const found: { story: string; index: number; payload: Record<string, unknown> }[] = [];
  let index = 0;
  for (const segment of engine.geometryReader.storySegments(story)) {
    if (segment.kind === 'embed' && segment.embedKind === kind) found.push({ story, index, payload: segment.payload });
    index += units(segment);
  }
  return found;
}

test('no live table names a hidden or missing cell story after any step of random two-peer table edits', () => {
  return (async () => {
    (selection as { setCellLocationShadowCheck?: (on: boolean) => void }).setCellLocationShadowCheck?.(true);
    const report: string[] = [];
    const violations: string[] = [];
    const kinds: Record<string, number> = {};
    let firstThrowLogged = -1;
    const RAW_OPS = new Set(['detachControl', 'detachTable', 'moveTable', 'shareCell', 'reattachControl']);
    const danglingLogged = new Set<string>();
    const familyTotals: Record<string, { danglingSteps: number; violations: number; projectionThrows: number; peerMissingSteps: number; invalidSeeds: number; peerSkipped: number; reused: number; built: number }> = {};
    const cacheCounts = () => (selection as { cellLocationCacheCounts?: () => { built: number; reused: number } }).cellLocationCacheCounts?.() ?? { built: 0, reused: 0 };
    for (const family of ['valid', 'adversarial'] as const) {
    const totals = familyTotals[family] = { danglingSteps: 0, violations: 0, projectionThrows: 0, peerMissingSteps: 0, invalidSeeds: 0, peerSkipped: 0, reused: 0, built: 0 };
    const familyStart = cacheCounts();
    for (const seed of SEEDS) {
      const next = random(seed);
      const pick = <T>(items: readonly T[]): T | undefined => (items.length ? items[Math.floor(next() * items.length)] : undefined);
      const arm = await openCommandEngine([], { bodyXml: body, clientId: 8300 + seed }).catch((error: unknown) => {
        console.log(`LOCAL-OPEN-FAILED seed=${seed} ${String((error as Error)?.stack ?? error).split('\n').slice(0, 8).join(' | ')}`);
        throw error;
      });
      let local = arm.engine;
      let legacy = withoutCellLocation(local);
      let invalidFrom: number | null = null;
      const projectionThrows = () => { try { createYrsPositionProjection(local.geometryReader, 'body'); return false; } catch (error) { if (isEngineTrap(error)) throw error; return true; } };
      local.beginUndoCapture();
      const openPeer = async (clientId: number): Promise<DocxWorkerCommandEngine | null> => {
        try {
          const opened = (await openCommandEngine([], { bodyXml: body, clientId, checkpoint: local.encodeState() })).engine;
          opened.beginUndoCapture();
          return opened;
        } catch (error) {
          if (isEngineTrap(error)) throw error;
          console.log(`PEER-UNAVAILABLE seed=${seed} client=${clientId} ${String((error as Error)?.stack ?? error).split('\n').slice(0, 6).join(' | ')}`);
          return null;
        }
      };
      let peer = await openPeer(8400 + seed);
      let engine = local;
      let remoteSteps = 0;
      let reseeds = 0;
      let keyMoves = 0;
      const history: string[] = [];
      let rebuildSameKey = 0;
      const lastBuildKey = new Map<string, string>();
      const syncToPeer = () => peer?.applyUpdate(local.encodeStateAsUpdate(peer.encodeStateVector()));
      const syncFromPeer = () => peer && local.applyUpdate(peer.encodeStateAsUpdate(local.encodeStateVector()));
      const answer = (target: DocxWorkerCommandEngine, story: string): unknown => {
        try { return resolveLiveCell(target, story); } catch (error) {
          if (isEngineTrap(error)) throw error;
          const message = String((error as Error)?.message ?? error);
          return message.startsWith('cell location cache differs') ? `shadow: ${message.slice(0, 200)}` : `throw: ${message.slice(0, 120)}`;
        }
      };
      const seen = new Set<string>();
      let checks = 0;
      let answered = 0;
      let refused = 0;
      let applied = 0;
      const ops: Record<string, number> = {};
      const liveStories = () => engine.geometryReader.storyIds();
      const tables = () => {
        const projection = createYrsPositionProjection(engine.geometryReader, 'body');
        return projection ? [...projection.allTables()] : [];
      };
      const textStories = () => liveStories().filter((story) => {
        try { return engine.paragraphs(story).length > 0; } catch { return false; }
      });
      const steps: Record<string, () => unknown> = {
        insertRow: () => { const t = pick(tables()); const c = t && pick(t.cells); return c && engine.insertRow({ story: t.story, tableIndex: t.tableIndex, row: c.row, column: c.column }, next() < 0.5 ? 'above' : 'below'); },
        insertColumn: () => { const t = pick(tables()); const c = t && pick(t.cells); return c && engine.insertColumn({ story: t.story, tableIndex: t.tableIndex, row: c.row, column: c.column }, next() < 0.5 ? 'left' : 'right'); },
        deleteRow: () => { const t = pick(tables()); const c = t && pick(t.cells); const at = c && { story: t.story, tableIndex: t.tableIndex, row: c.row, column: c.column }; return at && engine.deleteRow({ anchor: at, head: at }); },
        deleteColumn: () => { const t = pick(tables()); const c = t && pick(t.cells); const at = c && { story: t.story, tableIndex: t.tableIndex, row: c.row, column: c.column }; return at && engine.deleteColumn({ anchor: at, head: at }); },
        mergeCells: () => {
          const t = pick(tables()); const c = t && pick(t.cells); if (!t || !c) return undefined;
          const across = next() < 0.5;
          const head: YrsCellLoc = { story: t.story, tableIndex: t.tableIndex, row: c.row + (across ? 0 : 1), column: c.column + (across ? c.colspan : 0) };
          return engine.mergeCells({ anchor: { story: t.story, tableIndex: t.tableIndex, row: c.row, column: c.column }, head });
        },
        splitCell: () => { const t = pick(tables()); const c = t && pick(t.cells); return c && engine.splitCell({ story: t.story, tableIndex: t.tableIndex, row: c.row, column: c.column }, 1 + Math.floor(next() * 2), 1 + Math.floor(next() * 2)); },
        nestTable: () => {
          const story = pick(textStories().filter((id) => /:t\d+:r\d+c\d+/.test(id))); if (!story) return undefined;
          const paraId = engine.paragraphs(story)[0]!.paraId;
          return engine.insertTable({ story, paraId, offset: 0 }, 1 + Math.floor(next() * 2), 1 + Math.floor(next() * 2));
        },
        bodyTable: () => { const paraId = engine.paragraphs('body')[0]!.paraId; return engine.insertTable({ story: 'body', paraId, offset: 0 }, 2, 2); },
        deleteTable: () => { const t = pick(tables()); return t && engine.deleteTable({ story: t.story, tableIndex: t.tableIndex }); },
        detachControl: () => { const e = pick(embeds(engine, 'body', 'blockSdt')); return e && engine.applyRawOps('body', [{ op: 'delete', index: e.index, len: 1 }]); },
        detachTable: () => {
          const owner = pick(liveStories()); const e = owner ? pick(embeds(engine, owner, 'table')) : undefined;
          return e && engine.applyRawOps(e.story, [{ op: 'delete', index: e.index, len: 1 }]);
        },
        moveTable: () => {
          const owner = pick(liveStories()); const e = owner ? pick(embeds(engine, owner, 'table')) : undefined;
          const target = pick(textStories()); if (!e || !target || target === e.story) return undefined;
          engine.applyRawOps(e.story, [{ op: 'delete', index: e.index, len: 1 }]);
          return engine.applyRawOps(target, [{ op: 'insertEmbed', index: 0, kind: 'table', payload: e.payload }]);
        },
        shareCell: () => {
          const t = pick(tables()); const shared = t && pick(t.cells);
          const host = pick(textStories().filter((id) => shared && id !== shared.story && !id.startsWith(`${shared.story}:`) && !shared.story.startsWith(`${id}:`)));
          if (!shared || !host) return undefined;
          return engine.applyRawOps(host, [{ op: 'insertEmbed', index: 0, kind: 'table', payload: {
            tblPr: {}, grid: [1500], rows: [{ trPr: {}, cells: [{ tcPr: { rowspan: 1, colspan: 1 }, story: shared.story }] }],
          } }]);
        },
                reattachControl: () => {
          const story = pick([...seen].filter((id) => /^body:sdt\d+$/.test(id) && liveStories().includes(id))); if (!story) return undefined;
          if (embeds(engine, 'body', 'blockSdt').some((e) => e.payload.story === story)) return undefined;
          return engine.applyRawOps('body', [{ op: 'insertEmbed', index: 0, kind: 'blockSdt', payload: { story } }]);
        },
        cloneItem: () => { const item = engine.controlTargets().find((target) => target.control.controlType === 'repeatingSectionItem'); return item && engine.repeatControlAt(item.at, false); },
        removeItem: () => { const item = engine.controlTargets().find((target) => target.control.controlType === 'repeatingSectionItem'); return item && engine.repeatControlAt(item.at, true); },
        undo: () => engine.undo(),
        redo: () => engine.redo(),
      };
      const names = Object.keys(steps).filter((name) => family === 'adversarial' || !RAW_OPS.has(name));
      const danglingIn = (target: DocxWorkerCommandEngine): string[] => {
        const stories = target.geometryReader.storyIds();
        const missing: string[] = [];
        for (const story of stories) {
          let segments;
          try { segments = target.geometryReader.storySegments(story); } catch { continue; }
          for (const segment of segments) {
            if (segment.kind !== 'embed' || segment.embedKind !== 'table') continue;
            for (const row of (segment.payload.rows as { cells?: { story?: string }[] }[] | undefined) ?? []) {
              for (const named of row.cells ?? []) if (named.story && !stories.includes(named.story)) missing.push(named.story);
            }
          }
        }
        return missing;
      };
      const check = (step: number, op: string) => {
        const dangling = [...danglingIn(local), ...(peer ? danglingIn(peer) : [])];
        if (dangling.length) {
          totals.danglingSteps++;
          if (!danglingLogged.has(`${family}:${seed}`)) {
            danglingLogged.add(`${family}:${seed}`);
            console.log(`DANGLING ${family} seed=${seed} step=${step} op=${op} cells=${[...new Set(dangling)].join(',')} history=${history.slice(-8).join(' > ')}`);
          }
        }
        for (const story of local.geometryReader.storyIds()) seen.add(story);
        for (const story of peer?.geometryReader.storyIds() ?? []) seen.add(story);
        const keyOf = () => (local as { cellStructureKey?: () => string }).cellStructureKey?.() ?? '';
        let previousKey = keyOf();
        for (const story of [...seen, 'body:t99:r0c0', 'hf:rId9:t0:r0c0', 'custom:t0:r0c0', 'custom']) {
          const currentKey = keyOf();
          if (currentKey !== previousKey) { keyMoves++; if (keyMoves <= 3) console.log(`KEY-MOVED seed=${seed} step=${step} op=${op} before=${previousKey} after=${currentKey} at=${story}`); }
          previousKey = currentKey;
          checks++;
          const read = (local as { cellLocation?: (story: string) => unknown }).cellLocation;
          const fast = typeof read === 'function' ? read(story) ?? null : null;
          const projected = answer(legacy, story);
          const counts = (selection as { cellLocationCacheCounts?: () => { built: number } }).cellLocationCacheCounts;
          const builtBefore = counts?.().built ?? 0;
          const live = answer(local, story);
          if ((counts?.().built ?? 0) > builtBefore) {
            const root = selection.rootOf(story);
            if (lastBuildKey.get(root) === currentKey) { rebuildSameKey++; if (rebuildSameKey <= 3) console.log(`REBUILD-SAME-KEY seed=${seed} step=${step} root=${root} key=${currentKey} story=${story}`); }
            lastBuildKey.set(root, currentKey);
          }
          if (fast) answered++;
          if (typeof projected === 'string' && projected.startsWith('throw:')) {
            totals.projectionThrows++;
            if (family === 'valid' && firstThrowLogged !== seed) {
              firstThrowLogged = seed;
              console.log(`VALID-FIRST-THROW seed=${seed} step=${step} op=${op} story=${story} ${projected} history=${history.slice(-8).join(' > ')}`);
            }
          }
          const bad = (fast !== null && JSON.stringify(fast) !== JSON.stringify(projected)) || (projected === null && fast !== null) ||
            JSON.stringify(live) !== JSON.stringify(projected);
          if (bad) {
            totals.violations++;
            const located = (value: unknown) => value !== null && typeof value === 'object';
            const kind = typeof live === 'string' && live.startsWith('shadow:') ? 'stale-cache'
              : fast !== null && projected === null ? 'engine-answers-unreachable'
              : typeof projected === 'string' ? 'projection-throws'
              : fast === null && located(projected) && live === null ? 'engine-misses-cell'
              : fast !== null && projected !== null ? 'engine-differs'
              : 'resolved-differs';
            kinds[`${family}:${kind}`] = (kinds[`${family}:${kind}`] ?? 0) + 1;
            if (!kinds[`first:${family}:${kind}`]) { kinds[`first:${family}:${kind}`] = 1; console.log(`FIRST ${family}:${kind} seed=${seed} step=${step} op=${op} story=${story} engine=${JSON.stringify(fast)} projection=${JSON.stringify(projected)} resolved=${JSON.stringify(live)}`); }
          }
          if (bad && violations.length < 40) {
            violations.push(`${family} seed=${seed} step=${step} op=${op} story=${story} engine=${JSON.stringify(fast)} projection=${JSON.stringify(projected)} resolved=${JSON.stringify(live)}`);
          }
        }
      };
      check(0, 'open');
      for (let step = 1; step <= STEPS; step++) {
        const op = names[Math.floor(next() * names.length)]!;
        if (invalidFrom === null && family === 'adversarial' && projectionThrows()) {
          invalidFrom = step;
          totals.invalidSeeds++;
          console.log(`INVALID-RAW-STATE ${family} seed=${seed} from step ${step}`);
        }
        if (step === Math.floor(STEPS / 2)) {
          if (invalidFrom !== null) {
            totals.peerSkipped++;
            peer = null;
            console.log(`PEER-SKIPPED ${family} seed=${seed} step=${step}: invalid raw state since step ${invalidFrom}`);
          } else {
            try {
              const state = local.encodeState();
              const fresh = (await openCommandEngine(['Seed'], { clientId: 8500 + seed })).engine;
              fresh.openSourceDocx(await packageBytes(body), { state });
              local = fresh;
              legacy = withoutCellLocation(local);
              local.beginUndoCapture();
              reseeds++;
            } catch (error) {
              if (isEngineTrap(error)) throw error;
              console.log(`RESEED-FAILED ${family} seed=${seed} ${String((error as Error)?.stack ?? error).split('\n').slice(0, 6).join(' | ')}`);
            }
            peer = await openPeer(8600 + seed);
          }
          check(step, 'replace-document');
        }
        if (family === 'valid' && peer === null) totals.peerMissingSteps++;
        const remote = peer !== null && next() < 0.5;
        engine = remote && peer ? peer : local;
        if (remote) { syncToPeer(); remoteSteps++; }
        try {
          const result = steps[op]!();
          if (result === undefined || result === false) refused++; else { applied++; ops[op] = (ops[op] ?? 0) + 1; history.push(`${step}:${remote ? 'remote:' : ''}${op}`); }
        } catch (error) {
          if (isEngineTrap(error)) throw error;
          refused++;
        }
        engine.addUndoBoundary?.();
        if (remote) syncFromPeer();
        check(step, remote ? `remote:${op}` : op);
      }
      report.push(`INVARIANT ${family} seed=${seed} invalidFrom=${invalidFrom ?? 'none'} roots=${new Set([...seen].map(selection.rootOf)).size} steps=${STEPS} remote=${remoteSteps} reseeds=${reseeds} keyMovesWithinChecks=${keyMoves} rebuildsSameKey=${rebuildSameKey} applied=${applied} refused=${refused} storiesSeen=${seen.size} checks=${checks} engineAnswered=${answered} ops=${JSON.stringify(ops)}`);
      closeCommandEngines();
    }
    const familyEnd = cacheCounts();
    totals.reused = familyEnd.reused - familyStart.reused;
    totals.built = familyEnd.built - familyStart.built;
    }
    (selection as { setCellLocationShadowCheck?: (on: boolean) => void }).setCellLocationShadowCheck?.(false);
    const counts = (selection as { cellLocationCacheCounts?: () => unknown }).cellLocationCacheCounts?.();
    if (counts) console.log(`INVARIANT-CACHE ${JSON.stringify(counts)}`);
    for (const line of report) console.log(line);
    const total = Object.entries(kinds).filter(([key]) => !key.startsWith('first:')).reduce((sum, [, n]) => sum + n, 0);
    console.log(`INVARIANT-VIOLATIONS ${total} kinds=${JSON.stringify(Object.fromEntries(Object.entries(kinds).filter(([key]) => !key.startsWith('first:'))))}`);
    for (const [family, totals] of Object.entries(familyTotals)) {
      const rate = totals.reused + totals.built ? (totals.reused / (totals.reused + totals.built)).toFixed(3) : 'n/a';
      console.log(`INVARIANT-FAMILY ${family} ${JSON.stringify(totals)} reuseRate=${rate}`);
    }
    for (const line of violations.slice(0, 12)) console.log(`VIOLATION ${line}`);
    expect(violations).toEqual([]);
    expect(familyTotals.valid!.peerMissingSteps).toBe(0);
    expect(familyTotals.valid!.danglingSteps).toBe(0);
  })();
}, 600_000);
