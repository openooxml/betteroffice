import { afterEach, expect, mock, spyOn, test } from 'bun:test';
import type { YrsAuthor, YrsLoc, YrsStoryRange, YrsStorySegment } from './index';
import { preflightDocxCommand } from './residentCommandActions';
import { commandSelectionRange } from './residentCommandAnchors';
import type { DocxWorkerCommandEngine } from './residentCommandProtocol';
import { tableBoundaryDeleteTarget } from './tableBoundaryDelete';
import {
  TIMEOUT, closeCommandEngines, command, context, execute, openCommandEngine, selection, type CommandArm,
} from './residentCommandTestSupport';

afterEach(closeCommandEngines);

const SUGGESTING: YrsAuthor = { name: 'Writer', date: '2026-01-01T00:00:00Z' };

/** Reference delete preflight using story-wide paragraph spans. */
function storyScanDelete(
  engine: DocxWorkerCommandEngine, at: YrsLoc, backward: boolean, count: number, skipDeleted: boolean
): { range: YrsStoryRange; target?: YrsLoc } | 'unsupported-delete-target' {
  const reader = {
    paragraphSpans: (story: string) => engine.geometryReader.paragraphSpans(story),
    paragraphSegments: (story: string, indices: readonly number[]) => engine.paragraphSegments(story, indices),
  };
  const boundary = tableBoundaryDeleteTarget(reader, at, backward ? 'backward' : 'forward');
  if (boundary) return {
    target: boundary,
    range: { story: boundary.story, start: { paraId: boundary.paraId, offset: boundary.offset }, end: { paraId: boundary.paraId, offset: boundary.offset } },
  };
  const range = { story: at.story, start: { paraId: at.paraId, offset: at.offset }, end: { paraId: at.paraId, offset: at.offset } };
  const spans = engine.geometryReader.paragraphSpans(range.story);
  let index = spans.findIndex(({ paraId }) => paraId === range.start.paraId);
  if (index < 0) return 'unsupported-delete-target';
  let offset = range.start.offset;
  const paragraphAt = (index: number) => {
    const paragraph = { text: '', deleted: [] as Array<[number, number]> };
    for (const segment of engine.paragraphSegments(range.story, [index])[0] ?? []) {
      const text = segment.kind === 'text' ? segment.text : segment.kind === 'pilcrow' ? '\n' : '￼';
      if (skipDeleted && !(segment.kind === 'embed' && segment.embedKind === 'table') && (segment.attributes.del != null ||
        (segment.kind === 'pilcrow' && segment.properties.pPrDel != null))) {
        paragraph.deleted.push([paragraph.text.length, paragraph.text.length + text.length]);
      }
      paragraph.text += text;
    }
    return paragraph;
  };
  const tableBoundary = () => tableBoundaryDeleteTarget({
    paragraphSpans: story => story === range.story ? spans : engine.geometryReader.paragraphSpans(story),
    paragraphSegments: reader.paragraphSegments,
  }, { story: range.story, paraId: spans[index]!.paraId, offset }, backward ? 'backward' : 'forward');
  for (let left = count; left > 0; left -= 1) {
    const text = paragraphAt(index).text;
    const atBoundary = backward ? offset === 0 || text[offset - 1] === '￼' :
      offset === spans[index]!.length || text[offset] === '￼';
    if (atBoundary && tableBoundary()) break;
    if (backward) {
      if (offset === 0) {
        if (index === 0) break;
        offset = spans[--index]!.length;
      } else {
        offset -= 1;
        if (offset > 0 && /[\uDC00-\uDFFF]/u.test(text[offset]!) && /[\uD800-\uDBFF]/u.test(text[offset - 1]!)) offset -= 1;
      }
    } else {
      if (skipDeleted) {
        for (;;) {
          const deleted = paragraphAt(index).deleted.find(([start, end]) => start <= offset && offset < end);
          if (!deleted) break;
          offset = deleted[1];
          if (offset > spans[index]!.length) {
            if (index === spans.length - 1) {
              offset = spans[index]!.length;
              break;
            }
            index += 1;
            offset = 0;
          }
        }
        if (tableBoundary()) break;
      }
      if (offset === spans[index]!.length) {
        if (index === spans.length - 1) break;
        index += 1;
        offset = 0;
      } else {
        offset += paragraphAt(index).text.codePointAt(offset)! > 0xffff ? 2 : 1;
      }
    }
  }
  const target = { paraId: spans[index]!.paraId, offset };
  const expanded = { story: range.story, start: backward ? target : range.start, end: backward ? range.end : target };
  return commandSelectionRange({ geometryReader: engine.geometryReader }, {
    anchor: { story: range.story, ...expanded.start }, focus: { story: range.story, ...expanded.end },
  }) ? { range: expanded } : 'unsupported-delete-target';
}

const p = (id: string, inner: string, pPr = '') => `<w:p w14:paraId="${id}">${pPr}${inner}</w:p>`;
const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const stamp = (id: number) => `w:id="${id}" w:author="A" w:date="2026-01-01T00:00:00Z"`;
const table = (prefix: string, first: number) => '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  [0, 1].map(row => `<w:tr>${[0, 1].map(column => `<w:tc>${p((first + row * 2 + column).toString(16).padStart(8, '0'),
    r(`${prefix}${row}${column} \u{1F600}`))}</w:tc>`).join('')}</w:tr>`).join('') + '</w:tbl>';

/** Surrogate pairs and combining marks, a table at either end of a paragraph, tables back to back,
 * empty paragraphs, a field, inline and block content controls, a hyperlink, and tracked changes
 * including a deleted paragraph mark. */
const BOUNDARIES =
  p('00000001', r('Plain \u{1F600}\u{1F44D} é text')) +
  table('A', 0x101) +
  p('00000002', r('After a table')) +
  p('00000003', '') +
  table('B', 0x201) +
  table('C', 0x301) +
  p('00000004', r('Between tables')) +
  p('00000005', r('Field ') + '<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>' + r(' end')) +
  p('00000006', r('Control ') + '<w:sdt><w:sdtPr/><w:sdtContent>' + r('inline') + '</w:sdtContent></w:sdt>' + r(' after')) +
  '<w:sdt><w:sdtPr/><w:sdtContent>' + p('00000007', r('In a block control')) + '</w:sdtContent></w:sdt>' +
  p('00000008', r('Link ') + `<w:hyperlink r:id="rIdLink">${r('here')}</w:hyperlink>` + r(' done')) +
  p('00000009', r('Kept ') + `<w:ins ${stamp(1)}>${r('added')}</w:ins>` + `<w:del ${stamp(2)}><w:r><w:delText>gone</w:delText></w:r></w:del>` + r(' tail'),
    `<w:pPr><w:rPr><w:del ${stamp(3)}/></w:rPr></w:pPr>`) +
  p('0000000A', `<w:del ${stamp(4)}><w:r><w:delText>all deleted</w:delText></w:r></w:del>`) +
  p('0000000B', r('Last \u{1F600}'));

type DeleteCoverage = { carets: number; unresolved: number; cellTargets: number; refused: number; merges: number };

function expectDeletesMatchStoryScan(arm: CommandArm, label: string, seen: DeleteCoverage): void {
  const engine = arm.engine;
  for (const story of engine.geometryReader.storyIds()) {
    for (const span of engine.geometryReader.paragraphSpans(story)) {
      for (let offset = 0; offset <= span.length; offset++) {
        for (const backward of [true, false]) for (const count of [1, 2, 7]) for (const suggesting of [undefined, SUGGESTING]) {
          const kind = backward ? 'deleteBackward' : 'deleteForward';
          const intent = command(arm, kind, {
            selection: selection(span.paraId, offset, offset, story), count, ...(suggesting ? { suggesting } : {}),
          });
          const where = `${label} ${kind} x${count}${suggesting ? ' suggesting' : ''} ${story} ${span.paraId}@${offset}`;
          const result = preflightDocxCommand(engine, arm.state, intent, context(arm, intent));
          // A caret the selection check refuses never reaches the delete. The check without
          // `paragraphUnit` is the whole-story one, so it decides independently of the result.
          const caret = { story, paraId: span.paraId, offset };
          if (!commandSelectionRange({ geometryReader: engine.geometryReader }, { anchor: caret, focus: caret })) {
            expect(result.status === 'refused' ? result.reason : result.status, where).toBe('invalid-selection');
            seen.unresolved++;
            continue;
          }
          const expected = storyScanDelete(engine, { story, paraId: span.paraId, offset }, backward, count, !backward && suggesting !== undefined);
          if (expected === 'unsupported-delete-target') {
            expect(result.status === 'refused' ? result.reason : result.status, where).toBe(expected);
            seen.refused++;
          } else {
            expect(result.status, where).toBe('ready');
            if (result.status !== 'ready') continue;
            expect(result.prepared.range, where).toEqual(expected.range);
            if (expected.range.start.paraId !== expected.range.end.paraId) seen.merges++;
            if (expected.target?.story !== story && expected.target) seen.cellTargets++;
            if (expected.target) expect(result.prepared.selection, where).toEqual({ anchor: expected.target, focus: expected.target });
            else expect(result.prepared.selection, where).toEqual(selection(span.paraId, offset, offset, story));
          }
          seen.carets++;
        }
      }
    }
  }
}

test('backspace and delete read the caret paragraph by id and target what the story scan did', async () => {
  const seen = { carets: 0, unresolved: 0, cellTargets: 0, refused: 0, merges: 0 };
  const arm = await openCommandEngine([], { bodyXml: BOUNDARIES });
  const body = () => arm.engine.geometryReader.paragraphSpans('body');
  expect(body().map(span => span.paraId)).toContain('00000004');
  expectDeletesMatchStoryScan(arm, 'opened', seen);
  // Typing, a merge and a split each edit the story, so its paragraph index and spans are read anew.
  const plain = body()[0]!;
  expect(execute(arm, command(arm, 'insertText', { selection: selection(plain.paraId, 6), text: 'xy\u{1F600}' })).status).toBe('applied');
  expectDeletesMatchStoryScan(arm, 'after typing', seen);
  const last = body().at(-1)!;
  expect(execute(arm, command(arm, 'deleteBackward', { selection: selection(last.paraId, 0) })).status).toBe('applied');
  expectDeletesMatchStoryScan(arm, 'after a merge', seen);
  const link = body().find(span => span.paraId === '00000008')!;
  expect(execute(arm, command(arm, 'paragraphBreak', { selection: selection(link.paraId, 7) })).status).toBe('applied');
  expectDeletesMatchStoryScan(arm, 'after a split', seen);

  const duplicates = await openCommandEngine(['a', 'b', 'c', 'd']);
  const geometry = duplicates.engine.geometryReader;
  const { end } = geometry.locateParagraph('body', '00000003');
  duplicates.native.apply_seed_raw_ops('body', JSON.stringify([
    { op: 'setEmbedAttr', index: end, key: 'paraId', value: '00000002' },
  ]));
  expect(geometry.paragraphIdCount('body', '00000002')).toBe(2);
  const acrossDuplicates = command(duplicates, 'deleteForward', { selection: selection('00000001', 0), count: 7 });
  expect(preflightDocxCommand(duplicates.engine, duplicates.state, acrossDuplicates, context(duplicates, acrossDuplicates)))
    .toMatchObject({ status: 'refused', reason: 'unsupported-delete-target' });
  expectDeletesMatchStoryScan(duplicates, 'with duplicate paragraph ids', seen);
  console.log(`delete preflight parity: ${JSON.stringify(seen)}`);
  expect(seen.carets).toBeGreaterThan(4000);
  expect(seen.unresolved).toBeGreaterThan(0);
  expect(seen.cellTargets).toBeGreaterThan(0);
  expect(seen.refused).toBeGreaterThan(0);
  expect(seen.merges).toBeGreaterThan(100);
}, 120_000);

test('a direct forward table boundary check reads the caret paragraph segments once', () => {
  const at = { story: 'body', paraId: '00000001', offset: 1 };
  const paragraph: YrsStorySegment[] = [
    { kind: 'text', text: 'abc', attributes: {} },
    { kind: 'pilcrow', paraId: at.paraId, attributes: {}, properties: {} },
  ];
  const paragraphSegments = mock((_story: string, _indices: readonly number[]) => [paragraph]);
  expect(tableBoundaryDeleteTarget({
    paragraphSpans: () => { throw new Error('Unexpected story-wide paragraph read'); },
    paragraphSegments,
    paragraphUnit: () => 0,
  }, at, 'forward')).toBeNull();
  expect(paragraphSegments.mock.calls).toEqual([['body', [0]]]);
});

test('a backspace inside a paragraph reads that paragraph, not the story, at any document length', async () => {
  const work: Record<string, Array<{ spans: number; units: number }>> = {};
  for (const paragraphs of [40, 1200]) {
    const arm = await openCommandEngine(Array.from({ length: paragraphs }, (_, index) => `Paragraph ${index} of the long document \u{1F600}.`));
    const middle = arm.engine.geometryReader.paragraphSpans('body')[paragraphs / 2]!;
    // The keystroke before the backspace edits the story, so its spans are no longer cached.
    expect(execute(arm, command(arm, 'insertText', { selection: selection(middle.paraId, 10), text: 'x' })).status).toBe('applied');
    const spans = spyOn(arm.engine.geometryReader, 'paragraphSpans');
    const segments = spyOn(arm.engine, 'paragraphSegments');
    try {
      const cases = {
        'mid-paragraph backspace': ['deleteBackward', 11],
        'backspace over a surrogate pair': ['deleteBackward', middle.length],
        'backspace at the paragraph start': ['deleteBackward', 0],
        'mid-paragraph delete': ['deleteForward', 11],
      } as const;
      for (const [name, [kind, offset]] of Object.entries(cases)) {
        spans.mockClear();
        segments.mockClear();
        const intent = command(arm, kind, { selection: selection(middle.paraId, offset) });
        expect(preflightDocxCommand(arm.engine, arm.state, intent, context(arm, intent)).status, name).toBe('ready');
        (work[name] ??= []).push({
          spans: spans.mock.calls.filter(([story]) => story === 'body').length,
          units: segments.mock.calls.reduce((total, [, units]) => total + units.length, 0),
        });
      }
    } finally {
      spans.mockRestore();
      segments.mockRestore();
    }
  }
  expect(work).toEqual({
    'mid-paragraph backspace': [{ spans: 0, units: 1 }, { spans: 0, units: 1 }],
    'backspace over a surrogate pair': [{ spans: 0, units: 1 }, { spans: 0, units: 1 }],
    'backspace at the paragraph start': [{ spans: 0, units: 2 }, { spans: 0, units: 2 }],
    'mid-paragraph delete': [{ spans: 0, units: 1 }, { spans: 0, units: 1 }],
  });
}, TIMEOUT * 3);

test('counted suggesting forward delete visits paragraph segments a bounded number of times', async () => {
  const paraId = '00000001';
  const count = 8;
  const runs = Array.from({ length: 128 }, (_, index) => index % 2 === 0
    ? r('a') : '<w:r><w:rPr><w:b/></w:rPr><w:t>b</w:t></w:r>').join('');
  const arm = await openCommandEngine([], { bodyXml: p(paraId, runs) });
  const readSegments = arm.engine.paragraphSegments.bind(arm.engine);
  const segmentCount = readSegments('body', [0])[0]!.length;
  expect(segmentCount).toBeGreaterThan(count);
  let visits = 0;
  const segments = spyOn(arm.engine, 'paragraphSegments').mockImplementation((story, indices) =>
    readSegments(story, indices).map(paragraph => new Proxy(paragraph, {
      get(target, key, receiver) {
        if (typeof key === 'string' && /^\d+$/u.test(key)) visits += 1;
        return Reflect.get(target, key, receiver);
      },
    })));
  const spans = spyOn(arm.engine.geometryReader, 'paragraphSpans');
  try {
    const intent = command(arm, 'deleteForward', { selection: selection(paraId, 0), count, suggesting: SUGGESTING });
    expect(preflightDocxCommand(arm.engine, arm.state, intent, context(arm, intent))).toMatchObject({
      status: 'ready',
      prepared: { range: { story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: count } } },
    });
    expect(spans).not.toHaveBeenCalled();
    expect(segments).toHaveBeenCalledTimes(1);
    expect(visits).toBeLessThanOrEqual(2 * segmentCount + 1);
  } finally {
    spans.mockRestore();
    segments.mockRestore();
  }
}, TIMEOUT);
