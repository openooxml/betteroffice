import { expect } from 'bun:test';

import {
  STORY,
  applyStages,
  displayFrame,
  frameText,
  paragraphText,
  regionLayout,
  typing,
  typingTarget,
} from './context';
import type { DocxScenario } from './context';
import type { YrsEngineApplyProfile } from '../../../packages/docx/src/yrs';
import {
  applyFrameDeltaOwned,
  decodeFrameDelta,
} from '../../../packages/docx/src/layout/render/frameDelta';
import type { RetainedFrame } from '../../../packages/docx/src/layout/render/frameDelta';

const BURST = 'The quick brown fox jumps over the lazy dog, twice';
const SECOND = 'and again on the next line';

export const typingBurst: DocxScenario = {
  name: 'typing-burst',
  description:
    'Types fifty characters one keystroke at a time through the resident path, splits the paragraph, types twenty-six more, then backspaces them away, so per-keystroke stage cost and its drift over a burst are measured.',
  participants: ['web'],
  async run({ recorder, open }) {
    const editor = await recorder.loadAsync(() => open());
    const { session } = editor;
    const target = typingTarget(
      session,
      session.paragraphs(STORY),
      session.paragraphSpans(STORY)
    );
    const caret = { story: STORY, paraId: target.paraId, offset: target.end };
    recorder.op('setSelection', () => session.setSelection(caret));
    regionLayout(editor, recorder, 'layoutDocumentWithRegions:initial');
    let frame = displayFrame(
      session,
      recorder,
      'displayListFrame:initial',
      null
    );

    for (const character of BURST) {
      frame = typing(
        session,
        recorder,
        'applyInput:keystroke',
        character,
        frame
      );
    }
    expect(paragraphText(session, target.paraId)).toBe(
      `${target.text}${BURST}`
    );
    expect(frameText(frame)).toContain('lazy dog');

    const split = recorder.op('splitParagraph', () =>
      session.splitParagraph({
        story: STORY,
        paraId: target.paraId,
        offset: target.end + BURST.length,
      })
    );
    const second = split.secondParaId;
    recorder.op('setSelection:secondLine', () =>
      session.setSelection({ story: STORY, paraId: second, offset: 0 })
    );
    frame = displayFrame(
      session,
      recorder,
      'displayListFrame:afterSplit',
      frame
    );
    for (const character of SECOND) {
      frame = typing(
        session,
        recorder,
        'applyInput:secondLine',
        character,
        frame
      );
    }
    expect(paragraphText(session, second)).toBe(SECOND);

    for (let index = 0; index < SECOND.length; index += 1) {
      frame = deleting(session, recorder, 'applyDelete:backspace', frame);
    }
    expect(paragraphText(session, second)).toBe('');
    expect(paragraphText(session, target.paraId)).toBe(
      `${target.text}${BURST}`
    );
    expect(
      frameText(
        displayFrame(session, recorder, 'displayListFrame:afterDelete', frame)
      )
    ).toContain('lazy dog');
  },
};

function deleting(
  session: Parameters<typeof typing>[0],
  timer: Parameters<typeof typing>[1],
  op: string,
  previous: RetainedFrame
): RetainedFrame {
  let profile: YrsEngineApplyProfile | undefined;
  return timer.op(
    op,
    () => {
      const result = session.applyDeleteProfiled(
        'backward',
        previous.frameEpoch
      );
      profile = result.profile;
      return applyFrameDeltaOwned(previous, decodeFrameDelta(result.frame));
    },
    () => applyStages(profile!)
  );
}
