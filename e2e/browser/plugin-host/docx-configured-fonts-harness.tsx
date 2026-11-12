import { createRoot } from 'react-dom/client';
import { useEffect, useState } from 'react';
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
  ready: false,
  fontsReady: [] as number[],
  load(bytes: number[]) { open(Uint8Array.from(bytes).buffer); },
};
(window as unknown as { __configuredFonts: typeof probe }).__configuredFonts = probe;
const fontsLoaded = () => { probe.fontsReady.push(performance.now()); };

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  open = setBuffer;
  useEffect(() => { probe.ready = true; }, []);
  return buffer && <DocxEditor documentBuffer={buffer} onFontsLoaded={fontsLoaded} onError={(error) => probe.errors.push(error.message)} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
