import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import { DocxEditor, type DocxEditorProps, type DocxEditorRef } from '@betteroffice/docx-react';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import fontUrl from '../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf?url';
import '../../../packages/docx-react/src/styles/editor.css';

setGoogleFontsEnabled(false);
let open: (buffer: ArrayBuffer) => void = () => {};
let control: (comments: DocxEditorProps['comments']) => void = () => {};
let hosted: DocxEditorProps['comments'];
const probe = {
  editor: null as DocxEditorRef | null,
  errors: [] as string[],
  load(bytes: number[]) { open(Uint8Array.from(bytes).buffer); },
  setComments(comments: DocxEditorProps['comments']) { hosted = comments; control(comments); },
  pageInView(index: number) {
    const rect = document.querySelector(`.paged-editor__page[data-page-index="${index}"]`)?.getBoundingClientRect();
    return !!rect && rect.bottom > 0 && rect.top < window.innerHeight;
  },
  scrollToTop() {
    for (let element = document.querySelector('.paged-editor__page')?.parentElement; element; element = element.parentElement) element.scrollTop = 0;
  },
  async saved() { return [...new Uint8Array((await this.editor!.save())!)]; },
};
(window as unknown as { __workerComments: WorkerCommentsProbe }).__workerComments = probe;
const fonts = [{ family: 'Liberation Sans', src: fontUrl }];
const measurementFontProvider = { resolve: () => async () => (await fetch(fontUrl)).arrayBuffer() };

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  const [comments, setComments] = useState(hosted);
  open = setBuffer;
  control = setComments;
  return buffer && <DocxEditor ref={editor => { probe.editor = editor; }} documentBuffer={buffer} fonts={fonts} comments={comments}
    measurementFontProvider={measurementFontProvider} showToolbar={false} onError={error => probe.errors.push(error.message)} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
