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
  mirrorMutatedAt: 0,
  load(bytes: number[]) { open(Uint8Array.from(bytes).buffer); },
};
(window as unknown as { __configuredFonts: typeof probe }).__configuredFonts = probe;
const fontsLoaded = () => { probe.fontsReady.push(performance.now()); };
const MIRROR = '.canvas-page-mirror';
const touchesMirror = (node: Node) => node instanceof Element && (node.matches(MIRROR) || node.querySelector(MIRROR) !== null);
new MutationObserver((records) => {
  const at = performance.now();
  if (records.some((record) => {
    const target = record.target instanceof Element ? record.target : record.target.parentElement;
    return target?.closest(MIRROR) || [...record.addedNodes, ...record.removedNodes].some(touchesMirror);
  })) probe.mirrorMutatedAt = at;
}).observe(document.body, { subtree: true, childList: true, characterData: true, attributes: true });

function Harness() {
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  open = setBuffer;
  useEffect(() => { probe.ready = true; }, []);
  return buffer && <DocxEditor documentBuffer={buffer} onFontsLoaded={fontsLoaded} onError={(error) => probe.errors.push(error.message)} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
