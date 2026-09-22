import { expect } from 'bun:test';

import { FONT_FAMILY, editStages, firstStory, profiled, profiledLayout, storyText } from './context';
import type { PptxScenario } from './context';
import { PythonWorker, fromBase64, pythonMissing, toBase64 } from '../python';

const SLIDE = 0;
const MARKER = 'WEB-MARKER ';
const BOX = { x: 914_400, y: 4_114_800, width: 3_657_600, height: 914_400 };

const OPEN = `
def run(state, input, timed):
    with timed('import'):
        import betteroffice_pptx as bp
    with timed('open'):
        deck = bp.Presentation.open(base64.b64decode(input['bytes']))
    state['deck'] = deck
    with timed('read'):
        slides = [{'id': slide.id, 'shapes': len(slide.shapes), 'text': slide.text} for slide in deck.snapshot().slides]
    return {'slideCount': deck.slide_count, 'slideIds': deck.slide_ids, 'slides': slides}
`;

const EDIT = `
def run(state, input, timed):
    deck = state['deck']
    with timed('findStory'):
        story_id = None
        for slide in deck.snapshot().slides:
            for shape in slide.shapes:
                for story in shape.stories:
                    if input['marker'] in story.text:
                        story_id = story.id
        if story_id is None:
            raise AssertionError('the web marker is in no story')
    with timed('insertText'):
        deck.insert_text(story_id, 0, input['insert'])
    with timed('addTextBox'):
        box = deck.add_text_box(input['slideId'], name='python box', x=input['rect']['x'], y=input['rect']['y'], width=input['rect']['width'], height=input['rect']['height'], text=input['boxText'], font_size=16.0)
    with timed('insertSlide'):
        appended = deck.insert_slide(deck.slide_count)
    with timed('save'):
        data = deck.save()
    return {'storyId': story_id, 'shapeId': box.shape_id, 'slideId': appended.slide_id, 'slideCount': deck.slide_count, 'bytes': base64.b64encode(data).decode()}
`;

export const pythonRoundtrip: PptxScenario = {
  name: 'python-roundtrip',
  description: 'The web core edits a deck and saves it, the Python binding reopens those bytes, checks every slide it sees against the web snapshot, makes its own edits and saves, and the web core reopens and verifies the result.',
  participants: ['web', 'python'],
  requires: pythonMissing,
  async run({ recorder, open }) {
    const handle = recorder.load(() => open());
    const deck = recorder.op('snapshot', () => handle.snapshot());
    const slide = deck.slides[SLIDE];
    const story = firstStory(slide.shapes);
    const original = storyText(story);

    profiled(recorder, 'insertText:web', () => handle.insertTextProfiled(story.id, 0, MARKER), editStages);
    profiled(
      recorder,
      'addTextBox:web',
      () => handle.addTextBoxProfiled(slide.id, { name: 'web box', rect: BOX, text: 'Added by the web core', style: { fontSizePt: 18, fontFamily: FONT_FAMILY } }),
      editStages
    );
    const handed = recorder.op('save', () => handle.save());
    const edited = recorder.op('snapshot:handed', () => handle.snapshot());

    const worker = new PythonWorker(recorder);
    try {
      await worker.start();
      const opened = await worker.call<{ slideCount: number; slideIds: string[]; slides: { id: string; shapes: number; text: string }[] }>(
        'python:open',
        OPEN,
        { bytes: toBase64(handed) },
        { bytes: handed.byteLength }
      );
      expect(opened.slideCount).toBe(edited.slides.length);
      expect(opened.slideIds).toEqual(edited.slides.map((entry) => entry.id));
      expect(opened.slides[SLIDE].shapes).toBe(edited.slides[SLIDE].shapes.length);
      expect(opened.slides[SLIDE].text).toContain(MARKER.trim());

      const back = await worker.call<{ storyId: string; shapeId: string; slideId: string; slideCount: number; bytes: string }>('python:editAndSave', EDIT, {
        slideId: opened.slides[SLIDE].id,
        marker: MARKER.trim(),
        insert: 'PY-MARKER ',
        boxText: 'Added by the python binding',
        rect: { ...BOX, y: BOX.y + 914_400 },
      });
      expect(back.storyId).toBe(story.id);
      expect(back.slideCount).toBe(edited.slides.length + 1);

      const returned = fromBase64(back.bytes);
      const reopened = recorder.op('reopen:pythonBytes', () => open(returned));
      const final = recorder.op('snapshot:afterReopen', () => reopened.snapshot());
      expect(final.slides.length).toBe(edited.slides.length + 1);
      expect(final.slides[SLIDE].shapes.map((shape) => shape.name)).toContain('web box');
    expect(final.slides[SLIDE].shapes.map((shape) => shape.name)).toContain('python box');
      const text = final.slides[SLIDE].shapes.flatMap((shape) => shape.textStories.map(storyText)).join(' ');
      expect(text).toContain('PY-MARKER');
      expect(text).toContain(MARKER.trim());
      expect(text).toContain(original.slice(0, 8));
      expect(recorder.op('searchText:pythonBox', () => reopened.searchText('python binding')).length).toBeGreaterThan(0);
      profiledLayout(reopened, recorder, 'layoutSlide:afterReopen', SLIDE);
    } finally {
      worker.close();
    }
  },
};
