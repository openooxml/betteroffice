export const wasm = {
  docx: async () => {
    const [yrs, editor, layout, utils, render] = await Promise.all([import('@betteroffice/docx/yrs'), import('@betteroffice/docx/editor'), import('@betteroffice/docx/layout'), import('@betteroffice/docx/utils'), import('@betteroffice/docx/layout/render')]);
    return { yrs, editor, layout, utils, render };
  },
  xlsx: () => import('@betteroffice/xlsx'),
  pptx: () => import('@betteroffice/pptx'),
};
