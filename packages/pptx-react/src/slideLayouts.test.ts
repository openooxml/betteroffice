import { expect, test } from 'bun:test';
import type { SlideDisplayList } from '@betteroffice/pptx';
import { createT, en, zhCN, deepMerge, type LocaleStrings } from '@betteroffice/pptx-i18n';
import {
  defaultInsertLayoutPath,
  emptyPlaceholderGuides,
  isEmptyEditablePlaceholder,
  slideLayoutChoices,
} from './slideLayouts';

const tEn = createT(en);
const tZh = createT(deepMerge(en, zhCN) as LocaleStrings);

test('defaultInsertLayoutPath stays within the active slide master in multi-master decks', () => {
  const layouts = [
    { partPath: 'ppt/slideLayouts/slideLayout1.xml', name: 'Title Slide', layoutType: 'title', masterPartPath: 'ppt/slideMasters/slideMaster1.xml' },
    { partPath: 'ppt/slideLayouts/slideLayout2.xml', name: 'Title and Content', layoutType: 'obj', masterPartPath: 'ppt/slideMasters/slideMaster1.xml' },
    { partPath: 'ppt/slideLayouts/slideLayout3.xml', name: 'Title Slide', layoutType: 'title', masterPartPath: 'ppt/slideMasters/slideMaster2.xml' },
    { partPath: 'ppt/slideLayouts/slideLayout4.xml', name: 'Title and Content', layoutType: 'obj', masterPartPath: 'ppt/slideMasters/slideMaster2.xml' },
  ];

  expect(defaultInsertLayoutPath(layouts, 'ppt/slideLayouts/slideLayout3.xml')).toBe(
    'ppt/slideLayouts/slideLayout4.xml',
  );
  expect(defaultInsertLayoutPath(layouts, 'ppt/slideLayouts/slideLayout1.xml')).toBe(
    'ppt/slideLayouts/slideLayout2.xml',
  );
});

test('defaultInsertLayoutPath avoids picking title again when obj is absent', () => {
  const layouts = [
    { partPath: 'ppt/slideLayouts/slideLayout1.xml', name: 'Title Slide', layoutType: 'title', masterPartPath: 'm1' },
    { partPath: 'ppt/slideLayouts/slideLayout2.xml', name: 'Two Content', layoutType: 'twoObj', masterPartPath: 'm1' },
  ];
  expect(defaultInsertLayoutPath(layouts, 'ppt/slideLayouts/slideLayout1.xml')).toBe(
    'ppt/slideLayouts/slideLayout2.xml',
  );
});

test('slideLayoutChoices localizes standard layout names while preserving custom layout names', () => {
  const layouts = [
    { partPath: 'l1', name: 'Title and Content', layoutType: 'obj' },
    { partPath: 'l2', name: '季度汇报模板', layoutType: 'obj' },
    { partPath: 'l3', name: 'Executive Summary', layoutType: 'obj' },
    { partPath: 'l4', name: null, layoutType: 'twoObj' },
    { partPath: 'l5', name: null, layoutType: null },
  ];

  expect(slideLayoutChoices(layouts, undefined, tEn).map((choice) => choice.label)).toEqual([
    'Title and Content',
    '季度汇报模板',
    'Executive Summary',
    'Two Content',
    'Layout 5',
  ]);
  expect(slideLayoutChoices(layouts, undefined, tZh).map((choice) => choice.label)).toEqual([
    '标题和内容',
    '季度汇报模板',
    'Executive Summary',
    '两栏内容',
    '版式 5',
  ]);
});

test('emptyPlaceholderGuides filters out non-text placeholders and preserves anchor, alignment, and transform', () => {
  const list: SlideDisplayList = {
    contractVersion: 1,
    width: 960,
    height: 540,
    primitives: [
      {
        kind: 'textBox',
        objectId: 2,
        shapeId: 's-title',
        storyId: 'story:s-title:0',
        placeholderType: 'title',
        x: 60,
        y: 40,
        w: 840,
        h: 90,
        anchor: 'center',
        paragraphs: [{ align: 'left', level: 0, runs: [] }],
        lines: [],
        overflow: false,
        transform: { rotationDeg: 5 },
      },
      {
        kind: 'textBox',
        objectId: 3,
        shapeId: 's-footer',
        storyId: 'story:s-footer:0',
        placeholderType: 'ftr',
        x: 300,
        y: 500,
        w: 360,
        h: 24,
        anchor: 'center',
        paragraphs: [{ align: 'center', level: 0, runs: [] }],
        lines: [],
        overflow: false,
      },
      {
        kind: 'textBox',
        objectId: 4,
        shapeId: 's-date',
        storyId: 'story:s-date:0',
        placeholderType: 'dt',
        x: 40,
        y: 500,
        w: 160,
        h: 24,
        anchor: 'center',
        paragraphs: [{ align: 'left', level: 0, runs: [] }],
        lines: [],
        overflow: false,
      },
      {
        kind: 'textBox',
        objectId: 5,
        shapeId: 's-sldnum',
        storyId: 'story:s-sldnum:0',
        placeholderType: 'sldNum',
        x: 780,
        y: 500,
        w: 120,
        h: 24,
        anchor: 'center',
        paragraphs: [{ align: 'right', level: 0, runs: [] }],
        lines: [],
        overflow: false,
      },
    ],
  };

  const guides = emptyPlaceholderGuides(list, tEn);
  expect(guides).toHaveLength(1);
  expect(guides[0]).toMatchObject({
    shapeId: 's-title',
    placeholderType: 'title',
    prompt: 'Click to add title',
    anchor: 'center',
    align: 'left',
    transform: { rotationDeg: 5 },
  });
  expect(isEmptyEditablePlaceholder(list, 's-title', 'story:s-title:0')).toBe(true);
  expect(isEmptyEditablePlaceholder(list, 's-footer', 'story:s-footer:0')).toBe(false);

  const multiParagraphList: SlideDisplayList = {
    ...list,
    primitives: [
      {
        ...(list.primitives[0] as Extract<SlideDisplayList['primitives'][number], { kind: 'textBox' }>),
        paragraphs: [
          { align: 'left', level: 0, runs: [] },
          { align: 'left', level: 0, runs: [] },
        ],
      },
    ],
  };
  expect(emptyPlaceholderGuides(multiParagraphList, tEn)).toHaveLength(0);
  expect(isEmptyEditablePlaceholder(multiParagraphList, 's-title', 'story:s-title:0')).toBe(false);
});
