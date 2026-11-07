import { createRoot } from 'react-dom/client';
import { useState } from 'react';
import { XlsxEditor } from '@betteroffice/xlsx-react';

const surface = new URLSearchParams(location.search).get('surface') ?? 'peerless';
const route = (
  surface === 'in-thread' ? { legacySyncApi: true, experimentalWorkerOpen: false }
  : surface === 'legacy-worker' ? { legacySyncApi: true, experimentalWorkerOpen: true }
  : surface === 'viewer' ? { legacySyncApi: true, experimentalWorkerOpen: true, readOnly: true }
  : {}
) as Record<never, never>;
let open: (file: Uint8Array) => void = () => {};
const probe = {
  opened: [] as unknown[][],
  load(bytes: number[]) { open(Uint8Array.from(bytes)); },
};
window.open = (...args) => { probe.opened.push(args); return null; };
(window as unknown as { __mirrorLinks: typeof probe }).__mirrorLinks = probe;

function Harness() {
  const [state, setState] = useState<{ file: Uint8Array; key: number } | null>(null);
  open = (file) => setState((previous) => ({ file, key: (previous?.key ?? 0) + 1 }));
  return state && <XlsxEditor key={state.key} {...route} file={state.file} />;
}
createRoot(document.getElementById('root')!).render(<Harness />);
