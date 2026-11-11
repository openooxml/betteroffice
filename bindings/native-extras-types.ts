import type { XlsxFontFace } from '../packages/xlsx/src/wasm/pngTarget';

/** What a native binding has beyond the wasm engine it replaces: top-level functions and instance members per class. */
export interface NativeExtras {
  functions: object;
  members: Record<string, object>;
}

export interface XlsxNativeExtras {
  functions: {
    configureFontPacks(latin?: string, intl?: string, cjk?: string): boolean;
    fontPackDiagnosticsJson(): string;
    engineVersion(): string;
    engineHeapLiveBytes(): number;
    setHeapLimit(bytes: number): void;
  };
  members: {
    XlsxDocument: {
      renderPngTarget(targetJson: string, optionsJson: string, fonts?: ReadonlyArray<XlsxFontFace>): Uint8Array;
    };
  };
}

export interface PptxNativeExtras {
  functions: {
    configureFontPacks(latin?: string, intl?: string, cjk?: string): boolean;
    fontPackDiagnosticsJson(): string;
    engineVersion(): string;
  };
  members: {
    PptxRenderer: {
      registerDefaultFonts(document: unknown, slide: number): void;
      renderPngTarget(document: unknown, targetJson: string, optionsJson: string): Uint8Array;
    };
  };
}

export interface DocxNativeExtras {
  functions: {};
  members: {
    EditSession: {
      renderPngTarget(displayListJson: string, targetJson: string, optionsJson: string): Uint8Array;
    };
  };
}
