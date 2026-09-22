import { expect } from 'bun:test';

import { editStages, firstStory, profiled, profiledLayout, storyLines, storyText } from './context';
import type { PptxScenario } from './context';
import { PythonWorker, fromBase64, pythonMissing, toBase64 } from '../python';

const SLIDE = 0;

const OPEN = `
def run(state, input, timed):
    with timed('import'):
        import betteroffice_pptx as bp
    with timed('open'):
        deck = bp.Presentation.open_collaborative(base64.b64decode(input['bytes']), client_id=input['clientId'])
    state['deck'] = deck
    with timed('stateVector'):
        vector = deck.state_vector()
    return {'clientId': deck.client_id, 'slideCount': deck.slide_count, 'stateVector': base64.b64encode(vector).decode()}
`;

const APPLY = `
def run(state, input, timed):
    deck = state['deck']
    with timed('applyUpdate'):
        deck.apply_update(base64.b64decode(input['update']))
    with timed('read'):
        story = deck.story(input['storyId']).text
        slides = deck.slide_ids
    with timed('stateVector'):
        vector = deck.state_vector()
    return {'story': story, 'slideIds': slides, 'stateVector': base64.b64encode(vector).decode()}
`;

const EDIT = `
def run(state, input, timed):
    deck = state['deck']
    with timed('insertText'):
        deck.insert_text(input['storyId'], input['index'], input['text'])
    with timed('diff'):
        update = deck.diff(base64.b64decode(input['peerVector']))
    with timed('read'):
        story = deck.story(input['storyId']).text
    return {'update': base64.b64encode(update).decode(), 'story': story}
`;

export const pythonWebLiveCollab: PptxScenario = {
  name: 'python-web-live-collab',
  description: 'A web replica and a Python replica of the same deck exchange Yrs updates in both directions, including concurrent inserts at the same offset, until both SDKs read back the same story text and slide order.',
  participants: ['web', 'python'],
  requires: pythonMissing,
  async run(ctx) {
    const { recorder, bytes, open } = ctx;
    const web = recorder.as('web');
    const handle = web.load(() => open(bytes, 1));
    const deck = web.op('snapshot', () => handle.snapshot());
    const story = firstStory(deck.slides[SLIDE].shapes);
    const original = storyLines(story);

    const worker = new PythonWorker(recorder);
    try {
      await worker.start();
      const opened = await worker.call<{ clientId: number; slideCount: number; stateVector: string }>('python:openCollaborative', OPEN, { bytes: toBase64(bytes), clientId: 2 }, { bytes: bytes.byteLength });
      expect(opened.clientId).toBe(2);
      expect(opened.slideCount).toBe(deck.slides.length);

      profiled(web, 'insertText:web', () => handle.insertTextProfiled(story.id, 0, 'WEB '), editStages);
      const appended = profiled(web, 'insertSlide:web', () => handle.insertSlideProfiled(deck.slides.length), editStages);
      let update = web.op('encodeDiff:toPython', () => handle.encodeDiff(fromBase64(opened.stateVector)));
      let mirrored = await worker.call<{ story: string; slideIds: string[]; stateVector: string }>(
        'python:applyUpdate',
        APPLY,
        { update: toBase64(update), storyId: story.id },
        { bytes: update.byteLength }
      );
      expect(mirrored.story).toBe(`WEB ${original}`);
      expect(mirrored.slideIds).toContain(appended.slideId);

      const back = await worker.call<{ update: string; story: string }>('python:insertAndDiff', EDIT, {
        storyId: story.id,
        index: 0,
        text: 'PY ',
        peerVector: toBase64(handle.encodeStateVector()),
      });
      const fromPython = fromBase64(back.update);
      web.op('applyUpdate:fromPython', () => handle.applyUpdate(fromPython), undefined, { bytes: fromPython.byteLength });
      expect(storyLines(web.op('story:afterPython', () => handle.story(story.id)))).toBe(back.story);
      expect(back.story).toBe(`PY WEB ${original}`);

      profiled(web, 'insertText:concurrent', () => handle.insertTextProfiled(story.id, 0, 'w'), editStages);
      const concurrent = await worker.call<{ update: string; story: string }>('python:concurrentInsert', EDIT, {
        storyId: story.id,
        index: 0,
        text: 'p',
        peerVector: toBase64(handle.encodeStateVector()),
      });
      const contested = fromBase64(concurrent.update);
      web.op('applyUpdate:concurrent', () => handle.applyUpdate(contested), undefined, { bytes: contested.byteLength });
      update = web.op('encodeDiff:catchUp', () => handle.encodeDiff(fromBase64(mirrored.stateVector)));
      mirrored = await worker.call<{ story: string; slideIds: string[]; stateVector: string }>(
        'python:applyUpdate',
        APPLY,
        { update: toBase64(update), storyId: story.id },
        { bytes: update.byteLength }
      );

      const merged = storyLines(handle.story(story.id));
      expect(mirrored.story).toBe(merged);
      expect(merged).toContain('w');
      expect(merged).toContain('p');
      expect(merged.endsWith(`PY WEB ${original}`)).toBe(true);
      expect(mirrored.slideIds).toEqual(web.op('snapshot:final', () => handle.snapshot()).slides.map((entry) => entry.id));
      profiledLayout(handle, web, 'layoutSlide:converged', SLIDE);
    } finally {
      worker.close();
    }
  },
};
