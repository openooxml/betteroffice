import React from 'react';
import { createRoot } from 'react-dom/client';
import { DocxEditor } from '@betteroffice/docx-react';
import { setGoogleFontsEnabled } from '@betteroffice/docx/utils';
import '@betteroffice/docx-react/styles.css';
import { createFontProvider } from '../../packages/fonts/src/cdn';

const provider = createFontProvider();
setGoogleFontsEnabled(false);
const fontLoads: {
  kind: string;
  family: string;
  ok: boolean;
  error?: string;
}[] = [];
const track =
  (kind: string, resolve: (...args: any[]) => (() => Promise<ArrayBuffer>) | undefined) =>
  (...args: any[]) => {
    const load = resolve(...args);
    if (!load) return undefined;
    return async () => {
      try {
        const bytes = await load();
        fontLoads.push({ kind, family: String(args[0]), ok: true });
        return bytes;
      } catch (error) {
        fontLoads.push({
          kind,
          family: String(args[0]),
          ok: false,
          error: String(error),
        });
        throw error;
      }
    };
  };
const measuredFonts = {
  resolve: track('family', provider.resolve),
  resolveScriptFallback: track('script', provider.resolveScriptFallback),
  resolveLastResort: track('fallback', provider.resolveLastResort),
};
let editor: React.RefObject<any>;
let root: ReturnType<typeof createRoot>;
const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
const api = window as any;
api.oracleInit = async (input: number[], fonts: boolean) => {
  fontLoads.length = 0;
  const bytes = new Uint8Array(input).buffer;
  root?.unmount();
  editor = React.createRef();
  api.errors = [];
  root = createRoot(document.getElementById('host')!);
  api.stage = 'editor';
  root.render(
    <DocxEditor
      ref={editor}
      documentBuffer={bytes}
      readOnly
      initialZoom={1}
      colorMode="light"
      showZoomControl={false}
      showOutlineButton={false}
      measurementFontProvider={fonts ? measuredFonts : undefined}
      onError={(error) => api.errors.push(String(error))}
    />
  );
  let last = 0;
  let stable = 0;
  for (;;) {
    if (api.errors.length > 0) throw new Error(api.errors.join('\n'));
    const pages = editor.current?.getTotalPages() ?? 0;
    const canvas = document.querySelector<HTMLCanvasElement>(
      'canvas[data-page-index="0"]'
    );
    if (pages > 0 && pages === last && canvas?.width && canvas?.height) {
      if (Date.now() - stable >= 3000) {
        api.stage = 'ready';
        return { pages, errors: api.errors, fontLoads };
      }
    } else {
      last = pages;
      stable = Date.now();
    }
    await sleep(100);
  }
};
api.oraclePage = async (index: number) => {
  editor.current.scrollToPage(index + 1);
  let previous = '';
  let equal = 0;
  for (;;) {
    const canvas = document.querySelector<HTMLCanvasElement>(
      `canvas[data-page-index="${index}"]`
    );
    if (canvas?.width && canvas?.height) {
      const output = document.createElement('canvas');
      output.width = canvas.width;
      output.height = canvas.height;
      const context = output.getContext('2d')!;
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, output.width, output.height);
      context.drawImage(canvas, 0, 0);
      const value = output.toDataURL('image/png');
      equal = value === previous ? equal + 1 : 0;
      if (equal >= 2) return value;
      previous = value;
    }
    await sleep(200);
  }
};
api.oracleReady = true;
