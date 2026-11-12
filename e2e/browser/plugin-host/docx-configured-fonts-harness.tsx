import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import { DocxEditor } from '@betteroffice/docx-react';
import { configureDefaultFonts } from '@betteroffice/docx/layout';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import * as fonts from '@betteroffice/fonts';
import '../../../packages/docx-react/src/styles/editor.css';

configureDefaultFonts({ fonts });
setGoogleFontsEnabled(false);
let open: (buffer: ArrayBuffer) => void = () => {};
const probe = {
  errors: [] as string[],
  load(bytes: number[]) { open(Uint8Array.from(bytes).buffer); },
};
(window as unknown as { __configuredFonts: typeof probe }).__configuredFonts = probe;

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  open = setBuffer;
  return buffer && <DocxEditor documentBuffer={buffer} onError={(error) => probe.errors.push(error.message)} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
