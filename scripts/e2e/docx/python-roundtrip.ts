import { expect } from 'bun:test';

import { STORY, paragraphText, save } from './context';
import type { YrsParagraph } from '../../../packages/docx/src/yrs';
import type { DocxScenario } from './context';
import { PythonWorker, fromBase64, pythonMissing, toBase64 } from '../python';

const MARKER = '[web marker]';
const REPLACEMENT = 'Replaced by the python binding';

const OPEN = `
def run(state, input, timed):
    with timed('import'):
        import betteroffice_docx as bd
    with timed('openOriginal'):
        original = bd.Document.open(base64.b64decode(input['original']))
    with timed('openEdited'):
        document = bd.Document.open(base64.b64decode(input['edited']))
    state['document'] = document
    with timed('read'):
        before = original.structure()
        after = document.structure()
        body = [p for section in document.sections() for p in section.paragraphs]
        addressable = [p for p in body if p.id is not None and len(p.runs) == 1 and p.text.strip() and input['marker'] not in p.text]
    return {
        'tablesBefore': before.body_tables,
        'tablesAfter': after.body_tables,
        'sections': after.sections,
        'text': document.text,
        'marked': [p.text for p in document.paragraphs() if input['marker'] in p.text],
        'paraId': addressable[0].id if addressable else None,
        'paraText': addressable[0].text if addressable else None,
    }
`;

const REPLACE = `
def run(state, input, timed):
    document = state['document']
    with timed('replaceText'):
        edit = document.replace_text(input['paraId'], input['text'])
    with timed('read'):
        text = document.paragraph(input['paraId']).text
    with timed('save'):
        data = document.save()
    return {'paraId': edit.para_id, 'text': text, 'bytes': base64.b64encode(data).decode()}
`;

export const pythonRoundtrip: DocxScenario = {
  name: 'python-roundtrip',
  description:
    'The web engine types a marker and inserts a table, saves, and the python binding reopens those bytes, confirms the extra table and the marker, then rewrites a single-run paragraph by its w14 id and saves it back for the web engine to reopen. Documents without authored paragraph ids take the branch that asserts the binding refuses an unknown id.',
  participants: ['web', 'python'],
  requires: pythonMissing,
  async run(ctx) {
    const { recorder, open } = ctx;
    const editor = await recorder.loadAsync(() => open());
    const { session } = editor;
    const spans = new Map(
      session.paragraphSpans(STORY).map((span) => [span.paraId, span.length])
    );
    const target =
      addressable(session.paragraphs(STORY)) ??
      session.paragraphs(STORY).find((paragraph) => /\S/.test(paragraph.text))!;
    recorder.op('insertText:marker', () =>
      session.insertText(
        {
          story: STORY,
          paraId: target.paraId,
          offset: spans.get(target.paraId)!,
        },
        ` ${MARKER}`
      )
    );
    const table = recorder.op('insertTable', () =>
      session.insertTable(
        { story: STORY, paraId: target.paraId, offset: 0 },
        2,
        2
      )
    );
    expect(table.createdStoryIds).toHaveLength(4);
    expect(paragraphText(session, target.paraId)).toContain(MARKER);

    const handed = await save(editor, recorder);

    const worker = new PythonWorker(recorder);
    try {
      await worker.start();
      const opened = await worker.call<{
        tablesBefore: number;
        tablesAfter: number;
        sections: number;
        text: string;
        marked: string[];
        paraId: string | null;
        paraText: string | null;
      }>(
        'python:open',
        OPEN,
        {
          original: toBase64(ctx.bytes),
          edited: toBase64(handed),
          marker: MARKER,
        },
        { bytes: handed.byteLength }
      );
      expect(opened.tablesAfter).toBe(opened.tablesBefore + 1);
      expect(opened.sections).toBeGreaterThan(0);
      expect(opened.text).toContain(MARKER);
      expect(opened.marked).toHaveLength(1);
      if (opened.paraId === null) {
        expect(opened.marked[0]).toContain(MARKER);
        await expect(
          worker.call('python:replaceUnknown', REPLACE, {
            paraId: 'FFFFFFFF',
            text: REPLACEMENT,
          })
        ).rejects.toThrow();
        return;
      }

      const occurrences = (session_: typeof session) =>
        session_
          .storyIds()
          .flatMap((storyId) => session_.paragraphs(storyId))
          .filter((paragraph) => paragraph.text === opened.paraText).length;
      const before = occurrences(session);
      expect(before).toBeGreaterThan(0);
      const replaced = await worker.call<{
        paraId: string;
        text: string;
        bytes: string;
      }>('python:replaceAndSave', REPLACE, {
        paraId: opened.paraId,
        text: REPLACEMENT,
      });
      expect(replaced.paraId).toBe(opened.paraId);
      expect(replaced.text).toBe(REPLACEMENT);

      const returned = fromBase64(replaced.bytes);
      const reopened = await recorder.opAsync('reopen:pythonBytes', () =>
        open(returned)
      );
      const text = recorder
        .op('paragraphs:afterReopen', () =>
          reopened.session
            .storyIds()
            .flatMap((storyId) => reopened.session.paragraphs(storyId))
        )
        .map((paragraph) => paragraph.text)
        .join('\n');
      expect(text).toContain(REPLACEMENT);
      expect(text).toContain(MARKER);
      expect(occurrences(reopened.session)).toBe(before - 1);
      expect(
        recorder.op('storyIds:afterReopen', () => reopened.session.storyIds())
          .length
      ).toBe(session.storyIds().length);
    } finally {
      await worker.close();
    }
  },
};

/** A paragraph the python binding can name: its paraId round-trips as `w14:paraId`. */
function addressable(paragraphs: YrsParagraph[]): YrsParagraph | undefined {
  return paragraphs.find(
    (paragraph) =>
      /\S/.test(paragraph.text) && /^[0-9A-Fa-f]{8}$/.test(paragraph.paraId)
  );
}
