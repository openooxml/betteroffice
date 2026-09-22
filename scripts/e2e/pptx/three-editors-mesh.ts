import { expect } from 'bun:test';

import { editStages, exchange, fingerprint, firstStory, profiled, replicas, storyText } from './context';
import type { PptxScenario } from './context';

const SLIDE = 0;

export const threeEditorsMesh: PptxScenario = {
  name: 'three-editors-mesh-with-undo',
  description: 'Three replicas each type into a shared story and append a slide, exchange in a full mesh, then each undoes its own last change and exchanges again; every round must leave all three decks identical.',
  participants: ['web:a', 'web:b', 'web:c'],
  run(ctx) {
    const peers = replicas(ctx, ['web:a', 'web:b', 'web:c']);
    const [a] = peers;
    const deck = a.timer.op('snapshot', () => a.handle.snapshot());
    const story = firstStory(deck.slides[SLIDE].shapes);
    const original = storyText(story);
    const converged = (op: string) => {
      const reference = fingerprint(a.timer.op(op, () => a.handle.snapshot()));
      for (const peer of peers.slice(1)) expect(fingerprint(peer.timer.op(op, () => peer.handle.snapshot()))).toBe(reference);
      return reference;
    };

    const slides: string[] = [];
    peers.forEach((peer, index) => {
      profiled(peer.timer, 'insertText:own', () => peer.handle.insertTextProfiled(story.id, 0, `${index}`), editStages);
      const appended = profiled(peer.timer, 'insertSlide:own', () => peer.handle.insertSlideProfiled(peer.handle.snapshot().slides.length), editStages);
      slides.push(appended.slideId);
    });

    expect(exchange(peers)).toBeGreaterThan(0);
    const meshed = converged('snapshot:meshed');
    for (const id of slides) expect(meshed).toContain(id);
    const typed = storyText(a.handle.story(story.id));
    expect(typed.endsWith(original)).toBe(true);
    expect(typed.length).toBe(original.length + peers.length);

    for (const peer of peers) {
      expect(peer.timer.op('canUndo', () => peer.handle.canUndo())).toBe(true);
      profiled(peer.timer, 'undo:own', () => peer.handle.undoProfiled(), (profile) => ({ undo: profile.undoMs, snapshot: profile.snapshotMs, serialize: profile.serializeMs }));
    }
    exchange(peers, 'applyUpdate:afterUndo');
    const afterUndo = converged('snapshot:afterUndo');
    expect(afterUndo).not.toBe(meshed);
    for (const id of slides) expect(afterUndo).not.toContain(id);
    expect(storyText(a.handle.story(story.id))).toBe(typed);

    expect(a.timer.op('redo:own', () => a.handle.redo()).applied).toBe(true);
    exchange(peers, 'applyUpdate:afterRedo');
    const afterRedo = converged('snapshot:afterRedo');
    expect(afterRedo).toContain(slides[0]);
    expect([...peers[1].handle.encodeStateVector()].join()).toBe([...peers[2].handle.encodeStateVector()].join());
  },
};
