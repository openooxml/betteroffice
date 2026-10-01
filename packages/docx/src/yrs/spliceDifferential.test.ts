import { afterEach, beforeAll, describe, expect, it } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';

import { unzipContainer } from '../docx/wasm';
import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession, type YrsSession } from './index';
import { captureSessionSave, writeSessionSave } from './saveYrsDocx';
import { yrsToDocument } from './yrsToDocument';

const ROOT = resolve(import.meta.dir, '../../../..');
const FIXTURES = [
  'crates/docx-edit/tests/fixtures',
  'crates/betteroffice-docx/tests/corpus/fixtures',
  'packages/docx/src/yrs/__fixtures__',
];
const STORY_PART = /^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/;
const TRIALS = Number(process.env.SPLICE_DIFF_TRIALS ?? 4);
const TIMEOUT = Number(process.env.SPLICE_DIFF_TIMEOUT_MS ?? 60_000);
const AUTHOR = { name: 'Differential', date: '2026-01-01T00:00:00Z' };
const decoder = new TextDecoder();

function documents(): string[] {
  const roots = process.env.SPLICE_DIFF_DOCS
    ? [resolve(process.env.SPLICE_DIFF_DOCS)]
    : FIXTURES.map((dir) => join(ROOT, dir));
  const found: string[] = [];
  const visit = (path: string) => {
    if (statSync(path).isDirectory()) {
      for (const entry of readdirSync(path).sort()) visit(join(path, entry));
    } else if (path.endsWith('.docx')) {
      found.push(path);
    }
  };
  for (const root of roots) visit(root);
  return found;
}

function random(seed: string): () => number {
  let state = 2166136261;
  for (const character of seed) state = Math.imul(state ^ character.charCodeAt(0), 16777619);
  return () => {
    state = (state + 0x6d2b79f5) | 0;
    let value = Math.imul(state ^ (state >>> 15), 1 | state);
    value = (value + Math.imul(value ^ (value >>> 7), 61 | value)) ^ value;
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

function parts(bytes: Uint8Array | ArrayBuffer): Record<string, string> {
  const entries = unzipContainer(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  return Object.fromEntries(
    Object.entries(entries).map(([name, data]) => [name, decoder.decode(data)])
  );
}

/**
 * The `w:p` elements not nested in another `w:p`, the XML between them, and whether each sits in
 * a vertically merged continuation cell, whose content Word and the model do not read.
 */
function paragraphs(xml: string): { spans: string[]; gaps: string[]; continued: boolean[] } {
  const tag =
    /<w:tcPr\b(?:[^>"'/]|"[^"]*"|'[^']*')*>([\s\S]*?)<\/w:tcPr>|<(\/?)w:tc(?=[\s>/])((?:[^>"']|"[^"]*"|'[^']*')*)>|<(\/?)w:p(?=[\s>/])((?:[^>"']|"[^"]*"|'[^']*')*)>/g;
  const spans: string[] = [];
  const gaps: string[] = [];
  const continued: boolean[] = [];
  const cells: boolean[] = [];
  let depth = 0;
  let start = 0;
  let cursor = 0;
  for (const found of xml.matchAll(tag)) {
    if (found[1] !== undefined) {
      if (cells.length > 0) {
        cells[cells.length - 1] = /<w:vMerge\b(?![^>]*\sw:val="restart")/.test(found[1]);
      }
      continue;
    }
    if (found[3] !== undefined) {
      if (found[2]) cells.pop();
      else if (!found[3].trimEnd().endsWith('/')) cells.push(false);
      continue;
    }
    const match = [found[0], found[4], found[5]] as const;
    const at = found.index!;
    if (depth === 0 && !match[1]) continued.push(cells.includes(true));
    if (match[1]) {
      depth -= 1;
      if (depth === 0) {
        spans.push(xml.slice(start, at + match[0].length));
        cursor = at + match[0].length;
      }
    } else if (match[2]!.trimEnd().endsWith('/')) {
      if (depth === 0) {
        gaps.push(xml.slice(cursor, at));
        spans.push(match[0]);
        cursor = at + match[0].length;
      }
    } else {
      if (depth === 0) {
        gaps.push(xml.slice(cursor, at));
        start = at;
      }
      depth += 1;
    }
  }
  gaps.push(xml.slice(cursor));
  return { spans, gaps, continued };
}

/** Each paragraph's field group: the index of the first paragraph a complex field joins it to. */
function fieldGroups(spans: string[]): number[] {
  let depth = 0;
  let first = 0;
  return spans.map((span, index) => {
    if (depth === 0) first = index;
    for (const [, type] of span.matchAll(/<w:fldChar\b[^>]*\sw:fldCharType="(begin|end)"/g)) {
      depth = Math.max(0, depth + (type === 'begin' ? 1 : -1));
    }
    return first;
  });
}

function withoutIds(paragraph: string): string {
  return paragraph.replace(/^<w:p\b[^>]*>/, (start) =>
    start.replace(/\s+w14:(paraId|textId)="[^"]*"/g, '')
  );
}

function decode(text: string): string {
  return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g, (_, entity: string) => {
    if (entity.startsWith('#x')) return String.fromCodePoint(Number.parseInt(entity.slice(2), 16));
    if (entity.startsWith('#')) return String.fromCodePoint(Number(entity.slice(1)));
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[entity]!;
  });
}

function text(paragraph: string): string {
  const content = paragraph
    .replace(/<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/g, '')
    .replace(/<w:(pPr|rPr)\b[^>]*\/>/g, '')
    .replace(/<w:(pPr|rPr)\b[^>]*>[\s\S]*?<\/w:\1>/g, '');
  return [
    ...content.matchAll(
      /<w:(t|delText)\b[^>]*>([^<]*)<\/w:\1>|<w:(tab|cr)\b[^>]*\/>|<w:br\b(?![^>]*\sw:type="(?:page|column)")[^>]*\/>/g
    ),
  ]
    .map((match) => (match[1] ? decode(match[2]!) : match[3] === 'tab' ? '\t' : '\n'))
    .join('');
}

function relationships(all: Record<string, string>, part: string): Map<string, string> {
  const rels = all[part.replace(/^word\//, 'word/_rels/') + '.rels'] ?? '';
  return new Map(
    [...rels.matchAll(/<Relationship\b[^>]*>/g)].map(([element]) => [
      / Id="([^"]*)"/.exec(element)?.[1] ?? '',
      / Target="([^"]*)"/.exec(element)?.[1] ?? '',
    ])
  );
}

const MARKER =
  /<w:(commentRangeStart|commentRangeEnd|commentReference|bookmarkStart|bookmarkEnd|ins|del|moveFrom|moveTo|moveFromRangeStart|moveFromRangeEnd|moveToRangeStart|moveToRangeEnd|fldChar|fldSimple|footnoteReference|endnoteReference|permStart|permEnd|hyperlink)\b((?:[^>"']|"[^"]*"|'[^']*')*)>/g;

/** The range, revision, field, note and link markers of `xml`, link targets resolved. */
function markers(xml: string, links: Map<string, string>): string[] {
  return [...xml.matchAll(MARKER)].map(([, name, attributes]) => {
    const value = (key: string) => new RegExp(`\\s${key}="([^"]*)"`).exec(attributes!)?.[1];
    const relationship = value('r:id');
    return [
      name,
      value('w:id'),
      value('w:name'),
      value('w:author'),
      value('w:fldCharType'),
      value('w:instr') && decode(value('w:instr')!),
      value('w:anchor'),
      relationship && links.get(relationship),
    ].join('|');
  });
}

/**
 * `wanted` markers missing from `held`. The model holds a move as a deletion and an insertion, so
 * a held `moveFrom` or `moveTo` stands for a wanted `del` or `ins` with the same id and author.
 */
function missing(wanted: string[], held: string[]): string[] {
  const asRevision: Record<string, string> = { moveFrom: 'del', moveTo: 'ins' };
  held = held.map((marker) => {
    const [name, ...rest] = marker.split('|');
    return asRevision[name!] ? [asRevision[name!], ...rest].join('|') : marker;
  });
  const counts = new Map<string, number>();
  for (const marker of held) counts.set(marker, (counts.get(marker) ?? 0) + 1);
  return wanted.filter((marker) => {
    const count = counts.get(marker) ?? 0;
    counts.set(marker, count - 1);
    return count <= 0;
  });
}

const RANGE =
  /<w:(bookmark|commentRange|moveFromRange|moveToRange|perm)(Start|End)\b(?:[^>"']|"[^"]*"|'[^']*')*?\sw:id="([^"]*)"/g;

/** Range ids of `xml` whose start or end is missing. */
function unpaired(xml: string): Set<string> {
  const open = new Map<string, number>();
  for (const [, kind, edge, id] of xml.matchAll(RANGE)) {
    const key = `${kind}:${id}`;
    open.set(key, (open.get(key) ?? 0) + (edge === 'Start' ? 1 : -1));
  }
  return new Set([...open].filter(([, count]) => count !== 0).map(([key]) => key));
}

const MOVE_RANGE =
  /<w:(moveFromRangeStart|moveToRangeStart)\b(?:[^>"']|"[^"]*"|'[^']*')*?\sw:name="([^"]*)"/g;

/** Move names of `xml` holding only one of their two ranges, and `(containers)` for a half-move. */
function unpairedMoves(xml: string): Set<string> {
  const halves = new Map<string, Set<string>>();
  for (const [, half, name] of xml.matchAll(MOVE_RANGE)) {
    if (!halves.has(name!)) halves.set(name!, new Set());
    halves.get(name!)!.add(half!);
  }
  const unpaired = new Set([...halves].filter(([, held]) => held.size !== 2).map(([name]) => name));
  const containers = (name: string) => (xml.match(new RegExp(`<w:${name}[\\s>]`, 'g')) ?? []).length;
  if ((containers('moveFrom') === 0) !== (containers('moveTo') === 0)) unpaired.add('(containers)');
  return unpaired;
}

const strayAmpersands = (xml: string) =>
  (xml.match(/&(?!(#x[0-9a-fA-F]+|#[0-9]+|[A-Za-z][\w.:-]*);)/g) ?? []).length;

function edit(session: YrsSession, next: () => number): string[] {
  const pick = <T>(items: readonly T[]) => items[Math.floor(next() * items.length)]!;
  const candidates = session
    .storyIds()
    .flatMap((story) =>
      session
        .paragraphs(story)
        .filter((paragraph) => paragraph.text.length > 1)
        .map((paragraph) => ({ story, paragraph }))
    );
  if (candidates.length === 0) return [];
  const body = candidates.filter(({ story }) => story === 'body');
  const pool = body.length > 0 && next() < 0.7 ? body : candidates;
  const structural = next() < 0.15;
  const count = structural || next() < 0.5 ? 1 : 2 + Math.floor(next() * 4);
  const targets = new Map<string, (typeof candidates)[number]>();
  for (let attempt = 0; attempt < count * 4 && targets.size < count; attempt += 1) {
    const target = pick(pool);
    targets.set(`${target.story}\u0000${target.paragraph.paraId}`, target);
  }
  const applied: string[] = [];
  for (const { story, paragraph } of targets.values()) {
    const length = paragraph.text.length;
    const offset = Math.floor(next() * length);
    const end = Math.min(length, offset + 1 + Math.floor(next() * 4));
    const range = {
      story,
      start: { paraId: paragraph.paraId, offset },
      end: { paraId: paragraph.paraId, offset: end },
    };
    const kind = structural
      ? pick(['split', 'merge'] as const)
      : pick(['insert', 'delete', 'replace', 'bold', 'suggestInsert', 'suggestDelete'] as const);
    try {
      if (kind === 'insert') session.insertText({ story, paraId: paragraph.paraId, offset }, 'Q9Z');
      else if (kind === 'delete') session.deleteRange(range);
      else if (kind === 'replace') session.replaceRange(range, 'Q9Z');
      else if (kind === 'bold') session.toggleMark(range, { type: 'bold' });
      else if (kind === 'suggestInsert')
        session.insertText({ story, paraId: paragraph.paraId, offset }, 'Q9Z', AUTHOR);
      else if (kind === 'suggestDelete') session.deleteRange(range, AUTHOR);
      else if (kind === 'split') session.splitParagraph({ story, paraId: paragraph.paraId, offset });
      else session.mergeParagraphs(story, paragraph.paraId);
      applied.push(`${kind}@${story}`);
    } catch {
      // An edit the session refuses here (a protected or final paragraph) is not a save case.
    }
  }
  return applied;
}

const sessions: YrsSession[] = [];

async function open(bytes: Uint8Array): Promise<YrsSession> {
  const session = await createYrsSession({ clientId: 7 });
  sessions.push(session);
  session.openDocx(bytes, true);
  return session;
}

function texts(session: YrsSession): string[] {
  return session
    .storyIds()
    .sort()
    .flatMap((story) => [`#${story}`, ...session.paragraphs(story).map(({ text }) => text)]);
}

beforeAll(() =>
  preloadEditWasm(
    new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')))
  )
);

afterEach(() => {
  for (const session of sessions.splice(0)) session.destroy();
});

describe('a spliced session save', () => {
  for (const file of documents()) {
    const name = relative(process.env.SPLICE_DIFF_DOCS ?? ROOT, file);
    it(
      `matches the whole-part save outside the paragraphs it keeps: ${name}`,
      async () => {
        const source = new Uint8Array(readFileSync(file));
        const next = random(name);
        const failures: string[] = [];
        const tally = { spliced: 0, whole: 0, kept: 0, rewritten: 0 };
        for (let trial = 0; trial < TRIALS; trial += 1) {
          const session = await open(source);
          const edits = edit(session, next);
          if (edits.length === 0) continue;
          const base = session.materializeDocx()!;
          const capture = captureSessionSave(session);
          const document = yrsToDocument(session, base);
          const options = { updateModifiedDate: false };
          const save = async (splices: boolean) =>
            (
              await writeSessionSave(
                session,
                document,
                capture,
                base.originalBuffer!,
                options,
                () => true,
                false,
                () => splices
              )
            ).bytes;
          const splicedBytes = await save(true);
          const wholeBytes = await save(false);
          const spliced = parts(splicedBytes);
          const whole = parts(wholeBytes);
          const original = parts(base.originalBuffer!);
          const fail = (message: string) => failures.push(`trial ${trial} [${edits}] ${message}`);
          for (const part of new Set([...Object.keys(spliced), ...Object.keys(whole)])) {
            if (spliced[part] === whole[part]) {
              if (STORY_PART.test(part)) tally.whole += 1;
              continue;
            }
            if (!STORY_PART.test(part) || spliced[part] === undefined || whole[part] === undefined) {
              fail(`${part} differs from the whole-part save`);
              continue;
            }
            tally.spliced += 1;
            const xml = spliced[part]!;
            const from = paragraphs(original[part] ?? '');
            const kept = paragraphs(xml);
            const written = paragraphs(whole[part]!);
            if (strayAmpersands(xml) > 0) fail(`${part} holds an ampersand the loader repairs`);
            if (kept.spans.length !== from.spans.length || written.spans.length !== from.spans.length) {
              fail(
                `${part} paragraphs: source ${from.spans.length}, spliced ${kept.spans.length}, whole ${written.spans.length}`
              );
              continue;
            }
            kept.gaps.forEach((gap, index) => {
              if (gap !== from.gaps[index]) fail(`${part} XML before paragraph ${index} changed`);
            });
            const links = { spliced: relationships(spliced, part), whole: relationships(whole, part) };
            const groups = fieldGroups(from.spans);
            const grouped = new Map<number, number[]>();
            groups.forEach((first, index) => {
              if (!grouped.has(first)) grouped.set(first, []);
              grouped.get(first)!.push(index);
            });
            const members = (first: number) => grouped.get(first)!;
            const source = kept.spans.map(
              (paragraph, index) => withoutIds(paragraph) === withoutIds(from.spans[index]!)
            );
            kept.spans.forEach((paragraph, index) => {
              if (!source[index] && paragraph === written.spans[index]) {
                tally.rewritten += 1;
                if (members(groups[index]!).some((member) => source[member])) {
                  fail(`${part} field group of paragraph ${index} is partly kept`);
                }
                return;
              }
              if (!source[index]) {
                fail(`${part} paragraph ${index} is neither its source nor the written XML`);
                return;
              }
              tally.kept += 1;
              if (groups[index] !== index || kept.continued[index]) return;
              const group = members(index);
              const join = (spans: string[]) => group.map((member) => spans[member]!).join('');
              if (text(join(kept.spans)) !== text(join(written.spans))) {
                fail(`${part} kept paragraph ${index} text differs from the model`);
              }
              const lost = missing(
                markers(join(written.spans), links.whole),
                markers(join(kept.spans), links.spliced)
              );
              if (lost.length > 0) fail(`${part} kept paragraph ${index} lacks ${lost.join(', ')}`);
            });
            const allowed = new Set([
              ...unpaired(whole[part]!),
              ...unpaired(original[part] ?? ''),
            ]);
            const broken = [...unpaired(xml)].filter((key) => !allowed.has(key));
            if (broken.length > 0) fail(`${part} unpaired ranges ${broken.join(', ')}`);
            const pairedBefore = unpairedMoves(original[part] ?? '');
            const lostMoves = [...unpairedMoves(xml)].filter((name) => !pairedBefore.has(name));
            if (lostMoves.length > 0) fail(`${part} moves without their other half ${lostMoves.join(', ')}`);
          }
          if (texts(await open(splicedBytes)).join('\n') !== texts(await open(wholeBytes)).join('\n')) {
            fail('reopened paragraph texts differ from the whole-part save');
          }
        }
        if (process.env.SPLICE_DIFF_REPORT) {
          console.log(`SPLICEDIFF ${name} ${JSON.stringify(tally)} failures ${failures.length}`);
        }
        expect(failures).toEqual([]);
      },
      TIMEOUT
    );
  }
});
