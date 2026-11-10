import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import type { DisplayPage, DisplayPrimitive } from './displayList';
import { buildMirrorPage, buildMirrorPageText } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const options = { labels: { page: 'Page 1', header: 'Header', footer: 'Footer', noteBacklink: 'Back' } };
const ADVANCE = 8;

function clusters(blockKey: string, start: number, text: string, x: number, baseline: number,
  separatorBefore?: string, separatorAfter?: string): DisplayPrimitive[] {
  return [...text].map((ch, index) => ({
    kind: 'glyphRun', fontId: 1, size: 16, color: '#000000', text: ch,
    glyphs: [{ id: 1, x: x + index * ADVANCE, y: baseline, cluster: 0, advance: ADVANCE }],
    blockKey, docStart: start + index, docEnd: start + index + 1, logicalOrder: start + index, bidiLevel: 0,
    ...(index === 0 && separatorBefore !== undefined ? { separatorBefore } : {}),
    ...(index === text.length - 1 && separatorAfter !== undefined ? { separatorAfter } : {}),
  }) as DisplayPrimitive);
}

const page = (primitives: DisplayPrimitive[]): DisplayPage => ({ pageIndex: 0, width: 600, height: 800, primitives });
const paragraphs = (mirror: HTMLElement) =>
  Array.from(mirror.querySelectorAll<HTMLElement>('[role="paragraph"]')).map((p) => p.textContent);

const cases: { name: string; primitives: DisplayPrimitive[]; read: string }[] = [
  {
    name: 'a line break',
    primitives: [...clusters('p1', 1, 'Special', 96, 140), ...clusters('p1', 9, 'Appropriation', 96, 160, '\n')],
    read: 'Special\nAppropriation',
  },
  {
    name: 'a tab',
    primitives: [...clusters('p1', 1, 'Tabbed', 96, 140), ...clusters('p1', 8, 'Word', 200, 140, '\t')],
    read: 'Tabbed\tWord',
  },
  {
    name: 'a double break',
    primitives: [...clusters('p1', 1, 'Double', 96, 140), ...clusters('p1', 9, 'Gap', 96, 180, '\n\n')],
    read: 'Double\n\nGap',
  },
  {
    name: 'a break that ends the paragraph',
    primitives: clusters('p1', 1, 'Endbreak', 96, 140, undefined, '\n'),
    read: 'Endbreak\n',
  },
  {
    name: 'a break that ends a column of a paragraph continued in the next column',
    primitives: [...clusters('p1', 1, 'Columnone', 96, 140, undefined, '\n'), ...clusters('p1', 11, 'Columntwo', 360, 140)],
    read: 'Columnone\nColumntwo',
  },
  {
    name: 'a break that opens the paragraph',
    primitives: clusters('p1', 2, 'Startbreak', 96, 160, '\n'),
    read: '\nStartbreak',
  },
];

for (const { name, primitives, read } of cases) {
  test(`${name} reads as its separator in the full mirror`, () => {
    expect(paragraphs(buildMirrorPage(page(primitives), options))).toEqual([read]);
  });

  test(`${name} reads as its separator in the text mirror`, () => {
    const mirror = buildMirrorPageText(page(primitives), options);
    expect(paragraphs(mirror)).toEqual([read]);
    const paragraph = mirror.querySelector('[role="paragraph"]')!;
    expect(Array.from(paragraph.childNodes).map((node) => node.nodeType)).toEqual([3]);
  });
}

test('the separator keeps touching runs on either side of it apart', () => {
  const primitives = [...clusters('p1', 1, 'Tab', 96, 140), ...clusters('p1', 4, 'Join', 120, 140, '\t')];
  const mirror = buildMirrorPage(page(primitives), options);
  expect(Array.from(mirror.querySelectorAll('.layout-run-text')).map((run) => run.textContent)).toEqual(['Tab', 'Join']);
  expect(paragraphs(mirror)).toEqual(['Tab\tJoin']);
});

test('soft wraps, hyphenation and CJK carry no separator and gain no node', () => {
  const primitives = [
    ...clusters('p1', 1, 'romeo ', 96, 140), ...clusters('p1', 7, 'sierra', 96, 160),
    ...clusters('p2', 20, 'state-of-', 96, 200), ...clusters('p2', 29, 'the-art', 96, 220),
    ...clusters('p3', 40, '日本語の', 96, 260), ...clusters('p3', 44, 'テキスト', 96, 280),
  ];
  for (const mirror of [buildMirrorPage(page(primitives), options), buildMirrorPageText(page(primitives), options)]) {
    expect(paragraphs(mirror)).toEqual(['romeo sierra', 'state-of-the-art', '日本語のテキスト']);
  }
  const full = buildMirrorPage(page(primitives), options);
  for (const paragraph of Array.from(full.querySelectorAll('[role="paragraph"]'))) {
    expect(Array.from(paragraph.childNodes).every((node) => node.nodeType === 1)).toBe(true);
  }
});
