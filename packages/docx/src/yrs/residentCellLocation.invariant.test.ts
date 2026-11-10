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

test('engine cell location is null or equal to the projection after every step of random table edits', () => {
  return (async () => {
    (selection as { setCellLocationShadowCheck?: (on: boolean) => void }).setCellLocationShadowCheck?.(true);
    const report: string[] = [];
    const violations: string[] = [];
    for (const seed of SEEDS) {
      const next = random(seed);
      const pick = <T>(items: readonly T[]): T | undefined => (items.length ? items[Math.floor(next() * items.length)] : undefined);
      const arm = await openCommandEngine([], { bodyXml: body, clientId: 8300 + seed }).catch((error: unknown) => {
        console.log(`LOCAL-OPEN-FAILED seed=${seed} ${String((error as Error)?.stack ?? error).split('\n').slice(0, 8).join(' | ')}`);
        throw error;
      });
      const local = arm.engine;
      const legacy = withoutCellLocation(local);
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
      const names = Object.keys(steps);
      const check = (step: number, op: string) => {
        for (const story of local.geometryReader.storyIds()) seen.add(story);
        for (const story of peer?.geometryReader.storyIds() ?? []) seen.add(story);
        for (const story of [...seen, 'body:t99:r0c0', 'hf:rId9:t0:r0c0', 'custom:t0:r0c0', 'custom']) {
          checks++;
          const read = (local as { cellLocation?: (story: string) => unknown }).cellLocation;
          const fast = typeof read === 'function' ? read(story) ?? null : null;
          const projected = answer(legacy, story);
          const live = answer(local, story);
          if (fast) answered++;
          const bad = (fast !== null && JSON.stringify(fast) !== JSON.stringify(projected)) || (projected === null && fast !== null) ||
            JSON.stringify(live) !== JSON.stringify(projected);
          if (bad && violations.length < 40) {
            violations.push(`seed=${seed} step=${step} op=${op} story=${story} engine=${JSON.stringify(fast)} projection=${JSON.stringify(projected)} resolved=${JSON.stringify(live)}`);
          }
        }
      };
      check(0, 'open');
      for (let step = 1; step <= STEPS; step++) {
        const op = names[Math.floor(next() * names.length)]!;
        if (step === Math.floor(STEPS / 2)) {
          try {
            const other = await openCommandEngine([], { bodyXml: body, clientId: 8500 + seed });
            local.loadState(other.engine.encodeState());
            reseeds++;
          } catch (error) {
            if (isEngineTrap(error)) throw error;
            console.log(`RESEED-FAILED seed=${seed} ${String((error as Error)?.stack ?? error).split('\n').slice(0, 6).join(' | ')}`);
          }
          peer = await openPeer(8600 + seed);
          check(step, 'reseed');
        }
        const remote = peer !== null && next() < 0.5;
        engine = remote && peer ? peer : local;
        if (remote) { syncToPeer(); remoteSteps++; }
        try {
          const result = steps[op]!();
          if (result === undefined || result === false) refused++; else { applied++; ops[op] = (ops[op] ?? 0) + 1; }
        } catch (error) {
          if (isEngineTrap(error)) throw error;
          refused++;
        }
        engine.addUndoBoundary?.();
        if (remote) syncFromPeer();
        check(step, remote ? `remote:${op}` : op);
      }
      report.push(`INVARIANT seed=${seed} steps=${STEPS} remote=${remoteSteps} reseeds=${reseeds} applied=${applied} refused=${refused} storiesSeen=${seen.size} checks=${checks} engineAnswered=${answered} ops=${JSON.stringify(ops)}`);
      closeCommandEngines();
    }
    (selection as { setCellLocationShadowCheck?: (on: boolean) => void }).setCellLocationShadowCheck?.(false);
    const counts = (selection as { cellLocationCacheCounts?: () => unknown }).cellLocationCacheCounts?.();
    if (counts) console.log(`INVARIANT-CACHE ${JSON.stringify(counts)}`);
    for (const line of report) console.log(line);
    console.log(`INVARIANT-VIOLATIONS ${violations.length}`);
    for (const line of violations.slice(0, 12)) console.log(`VIOLATION ${line}`);
    expect(violations).toEqual([]);
  })();
}, 600_000);
