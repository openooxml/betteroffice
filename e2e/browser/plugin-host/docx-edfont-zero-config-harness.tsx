import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import { DocxEditor } from '@betteroffice/docx-react';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import '../../../packages/docx-react/src/styles/editor.css';

setGoogleFontsEnabled(false);
const worker = new URLSearchParams(location.search).get('mode') !== 'legacy';
const route = (worker ? {} : { legacySyncApi: true, experimentalWorkerOpen: false }) as Record<never, never>;
let open: (buffer: ArrayBuffer) => void = () => {};
const probe = {
  errors: [] as string[],
  load(bytes: number[]) { open(Uint8Array.from(bytes).buffer); },
};
(window as unknown as { __edfontZeroConfig: typeof probe }).__edfontZeroConfig = probe;

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  open = setBuffer;
  return buffer && <DocxEditor {...route} documentBuffer={buffer} onError={(error) => probe.errors.push(error.message)} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
