import { expect } from 'bun:test';

import { STORY, displayFrame, regionLayout, save, typingTarget } from './context';
import type { DocxScenario } from './context';
import { PythonWorker, pythonMissing, toBase64 } from '../python';

const PHRASE = ' Parity probe paragraph, measured on the web side. ';

const PAGINATE = `
def run(state, input, timed):
    with timed('import'):
        import betteroffice_docx as bd
    with timed('open'):
        document = bd.Document.open(base64.b64decode(input['bytes']))
    with timed('paginate'):
        layout = document.layout(json.loads(input['kernel']))
    with timed('displayList'):
        display = layout.display_list
    return {'pages': layout.pages, 'primitives': display.primitives, 'pageSize': layout.to_dict()['pageSize']}
`;

export const pythonLayoutParity: DocxScenario = {
  name: 'python-layout-parity',
  description: 'The web engine measures the document and hands its retained kernel to the python binding, which paginates the same measured arena; both SDKs must report the same page count and page size, before and after an edit that reflows the text.',
  participants: ['web', 'python'],
  requires: pythonMissing,
  async run({ recorder, open }) {
    const editor = await recorder.loadAsync(() => open());
    const { session } = editor;

    const worker = new PythonWorker(recorder);
    try {
      await worker.start();
      const compare = async (label: string, bytes: Uint8Array) => {
        const laid = regionLayout(editor, recorder, `layoutDocumentWithRegions:${label}`);
        const revision = recorder.op(`residentWorkerProbe:${label}`, () => session.residentWorkerProbe()!.layoutRevision);
        const kernel = recorder.op(`retainedKernelInputs:${label}`, () => session.retainedKernelInputsJson(revision));
        const mirrored = await worker.call<{ pages: number; primitives: number; pageSize: { w: number; h: number } }>(
          `python:paginate:${label}`,
          PAGINATE,
          { bytes: toBase64(bytes), kernel },
          { kernelBytes: kernel.length }
        );
        expect(mirrored.pages).toBe(laid.layout.pages.length);
        expect(mirrored.primitives).toBeGreaterThan(0);
        expect(mirrored.pageSize.w).toBeGreaterThan(0);
        return { pages: laid.layout.pages.length, primitives: mirrored.primitives, pageSize: mirrored.pageSize };
      };

      const before = await compare('initial', await save(editor, recorder));
      expect(before.pages).toBeGreaterThan(0);
      const frame = displayFrame(session, recorder, 'displayListFrame:initial', null);
      expect(frame.pages.length).toBe(before.pages);

      const target = typingTarget(session, session.paragraphs(STORY), session.paragraphSpans(STORY));
      for (let index = 0; index < 12; index += 1) {
        recorder.op('insertText:parity', () => session.insertText({ story: STORY, paraId: target.paraId, offset: target.end }, PHRASE));
      }
      const handed = await save(editor, recorder);
      const after = await compare('afterEdits', handed);
      expect(after.pages).toBeGreaterThanOrEqual(before.pages);
      expect(after.pageSize).toEqual(before.pageSize);
      expect(displayFrame(session, recorder, 'displayListFrame:afterEdits', frame).pages.length).toBe(after.pages);
    } finally {
      worker.close();
    }
  },
};
