import type { PngFontFace } from './office-session/pngFonts';

export interface RequestedFontFace {
  family: string;
  bold: boolean;
  italic: boolean;
  scripts?: string[];
}

export interface DefaultFontFace {
  family: string;
  bold: boolean;
  italic: boolean;
  bytes: ArrayBuffer;
  kind?: 'primary' | 'coverage';
  face?: PngFontFace['face'];
  /** A lighter family's regular face drawn at its own weight whatever weight the run claims. */
  ownWeight?: boolean;
}

/** A default face the worker registered; its bytes reach the main thread only in the install request. */
export type DefaultFontLanding = Omit<DefaultFontFace, 'bytes' | 'face'>;

export interface DefaultFontPayload {
  defaultFonts?: DefaultFontLanding[];
}

export const defaultFontInstaller = 'pptxDefaultFontInstaller';

export type DefaultFontInstallation = {
  [defaultFontInstaller]?: (
    faces: readonly DefaultFontFace[], allowed?: (family: string) => boolean,
  ) => Promise<readonly DefaultFontFace[] | void>;
};

export interface DefaultFontInstallRequest {
  id: number;
  faces: DefaultFontFace[];
}
