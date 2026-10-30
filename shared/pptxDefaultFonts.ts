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
  /** A lighter family's regular face drawn at its own weight whatever weight the run claims. */
  ownWeight?: boolean;
}

export interface DefaultFontPayload {
  defaultFonts?: DefaultFontFace[];
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
