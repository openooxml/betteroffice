export interface DefaultFontFace {
  family: string;
  bold: boolean;
  italic: boolean;
  bytes: ArrayBuffer;
}

export interface DefaultFontPayload {
  defaultFonts?: DefaultFontFace[];
}

export const defaultFontInstaller = 'pptxDefaultFontInstaller';

export type DefaultFontInstallation = {
  [defaultFontInstaller]?: (faces: readonly DefaultFontFace[]) => Promise<void>;
};
