export const importFonts = (): Promise<typeof import('@betteroffice/fonts')> => import('@betteroffice/fonts');
export const importProgressiveFonts = (): Promise<typeof import('@betteroffice/fonts/progressive')> =>
  import('@betteroffice/fonts/progressive');
