import { expect, test } from 'bun:test';
import type { Document, Shape } from '../types/document';
import { recoverGroupedStorySeeds } from './groupedStorySeed';

function shape(text: string[]): Shape {
  return {
    type: 'shape', shapeType: 'rect', id: 'group', size: { width: 100, height: 100 },
    children: [{
      type: 'shape', shapeType: 'rect', id: 'box', size: { width: 100, height: 100 },
      textBody: { content: text.map((text, index) => ({
        type: 'paragraph', paraId: `0000000${index + 1}`,
        content: [{ type: 'run', content: [{ type: 'text', text }] }],
      })) },
    }],
  };
}

function document(group: Shape): Document {
  return { package: { document: { content: [{
    type: 'paragraph', content: [{ type: 'run', content: [{ type: 'shape', shape: group }] }],
  }] } } } as Document;
}

test('legacy grouped-story placeholders seed the exact retained paragraphs in source order', () => {
  const source = shape(['Alpha', 'Beta']);
  const sourceBytes = JSON.stringify(source);
  const placeholder = shape(['Alpha', 'Beta']);
  placeholder.children![0]!.textBody!.content = placeholder.children![0]!.textBody!.content.map(() => ({
    type: 'paragraph', formatting: {}, content: [],
  }));
  const target = document(placeholder);
  const restore = recoverGroupedStorySeeds(target, () => document(source));
  const host = target.package.document.content[0]!;
  if (host.type !== 'paragraph') throw new Error('host paragraph');
  const run = host.content[0]!;
  if (run.type !== 'run') throw new Error('host run');
  const drawing = run.content[0]!;
  if (drawing.type !== 'shape') throw new Error('host drawing');
  expect(JSON.stringify(drawing.shape)).toBe(sourceBytes);
  expect(JSON.stringify(source)).toBe(sourceBytes);
  restore();
  expect(drawing.shape).toBe(placeholder);
});

test('changed grouped-story text, paragraph order, and paragraph count are not rehydrated', () => {
  const source = shape(['Alpha', 'Beta']);
  for (const changed of [shape(['Edited', 'Beta']), shape(['Beta', 'Alpha']), shape(['Alpha'])]) {
    const target = document(changed);
    const before = JSON.stringify(target);
    let reads = 0;
    const restore = recoverGroupedStorySeeds(target, () => { reads++; return document(source); });
    expect(JSON.stringify(target)).toBe(before);
    expect(reads).toBe(0);
    restore();
  }
});

test('grouped-story placeholders with a different paragraph count cannot match the source', () => {
  const source = shape(['Alpha', 'Beta']);
  const placeholder = shape(['Alpha']);
  placeholder.children![0]!.textBody!.content = [{ type: 'paragraph', formatting: {}, content: [] }];
  const target = document(placeholder);
  const before = JSON.stringify(target);
  const restore = recoverGroupedStorySeeds(target, () => document(source));
  expect(JSON.stringify(target)).toBe(before);
  restore();
});

test('grouped-story placeholders do not choose between different retained source stories', () => {
  const first = document(shape(['Alpha', 'Beta']));
  first.package.document.content.push(...document(shape(['Edited', 'Beta'])).package.document.content);
  const placeholder = shape(['Alpha', 'Beta']);
  placeholder.children![0]!.textBody!.content = placeholder.children![0]!.textBody!.content.map(() => ({
    type: 'paragraph', formatting: {}, content: [],
  }));
  const target = document(placeholder);
  const before = JSON.stringify(target);
  const restore = recoverGroupedStorySeeds(target, () => first);
  expect(JSON.stringify(target)).toBe(before);
  restore();
});
