import { expect, test } from 'bun:test';
import { createPeerlessShellController } from './components/DocxEditor/peerless/shellController';
import { DocxEditorNotReadyError } from './index';

test('ref reads, edits and save before a document opens reject with the exported DocxEditorNotReadyError', async () => {
  const { ref } = createPeerlessShellController({ read: async () => ({} as never), focus() {}, zoom: 1 });
  const calls = [
    () => ref.readParagraphs({} as never),
    () => ref.findText({ text: 'x' } as never),
    () => ref.readSelectionInfo(),
    () => ref.applyEdits({ expectVersion: '0', steps: [] } as never),
    () => ref.save(),
  ];
  for (const call of calls) {
    let thrown: unknown;
    try { await call(); } catch (error) { thrown = error; }
    expect(thrown).toBeInstanceOf(DocxEditorNotReadyError);
    expect(thrown).toBeInstanceOf(Error);
    expect(thrown).toMatchObject({ name: 'DocxEditorNotReadyError', code: 'editor-not-ready', message: 'The editor is not ready' });
  }
});
