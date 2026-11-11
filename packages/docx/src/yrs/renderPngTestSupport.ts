import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';

export const ENGINES = [{ kind: 'wasm' }, { kind: 'native' }] as const;

export async function preloadEngines(): Promise<void> {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
}

let restore: (() => void) | undefined;

export function restoreBrowser(): void {
  restore?.();
  restore = undefined;
}

/** Canvas globals that record what a wasm-engine export draws; {@link restoreBrowser} removes them. */
export function stubBrowser() {
  restoreBrowser();
  const calls = { sizes: [] as [number, number][], fills: 0, encoded: 0, scales: [] as number[][], bitmaps: 0 };
  class Canvas {
    constructor(public width: number, public height: number) {}
    getContext() {
      if (calls.sizes.at(-1)?.[0] !== this.width || calls.sizes.at(-1)?.[1] !== this.height) calls.sizes.push([this.width, this.height]);
      return new Proxy({} as Record<string, unknown>, {
        get(target, property) {
          if (property === 'fill') return () => { calls.fills++; };
          if (property === 'scale') return (x: number, y: number) => { calls.scales.push([x, y]); };
          if (property === 'measureText') return (text: string) => ({ width: text.length });
          if (property === 'createLinearGradient' || property === 'createRadialGradient' || property === 'createPattern') return () => ({ addColorStop: () => undefined });
          return property in target ? target[property as string] : () => undefined;
        },
        set(target, property, value) { target[property as string] = value; return true; },
      });
    }
    convertToBlob(request: { type: string }) { calls.encoded++; return Promise.resolve(new Blob(['png'], { type: request.type })); }
  }
  class Path { moveTo() {} lineTo() {} quadraticCurveTo() {} bezierCurveTo() {} closePath() {} addPath() {} }
  const scope = globalThis as Record<string, unknown>;
  const saved = { OffscreenCanvas: scope.OffscreenCanvas, Path2D: scope.Path2D, Image: scope.Image, createImageBitmap: scope.createImageBitmap };
  Object.assign(scope, {
    OffscreenCanvas: Canvas, Path2D: Path, Image: undefined,
    createImageBitmap: () => { calls.bitmaps++; return Promise.resolve({ width: 4, height: 4, close: () => undefined }); },
  });
  restore = () => { Object.assign(scope, saved); };
  return calls;
}

export interface TextRun {
  family: string;
  text: string;
  bold?: boolean;
  rtl?: boolean;
}

const runXml = ({ family, text, bold, rtl }: TextRun) =>
  `<w:p>${rtl ? '<w:pPr><w:bidi/></w:pPr>' : ''}<w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:cs="${family}"/>${bold ? '<w:b/><w:bCs/>' : ''}${rtl ? '<w:rtl/>' : ''}</w:rPr><w:t>${text}</w:t></w:r></w:p>`;

/** A one-section document with a paragraph per run. */
export const runsDocument = (...runs: TextRun[]) => new Uint8Array(rezipPartsToArrayBuffer(new Map([
  ['[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
  ['_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
  ['word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${runs.map(runXml).join('')}<w:sectPr/></w:body></w:document>`)],
]) as PartsMap));

export const scriptDocument = () => runsDocument(
  { family: 'Calibri', text: 'Total 2024' },
  { family: 'Arial', text: 'مرحبا بالعالم', rtl: true },
  { family: 'Calibri', text: 'שלום עולם', rtl: true },
  { family: 'Brand Sans', text: 'Brand text' },
);
