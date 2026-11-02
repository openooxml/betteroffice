export type GenericFontFamily = 'serif' | 'sans-serif' | 'monospace' | 'cursive' | 'fantasy';

export interface WordFamilySubstitute {
  /** Look-alike family on a Google-compatible font CSS service, or null when none is close enough. */
  remote: string | null;
  generic: GenericFontFamily;
}

const sans = (remote: string | null): WordFamilySubstitute => ({ remote, generic: 'sans-serif' });
const serif = (remote: string | null): WordFamilySubstitute => ({ remote, generic: 'serif' });
const mono = (remote: string | null): WordFamilySubstitute => ({ remote, generic: 'monospace' });

// Our own table. The metric-compatible pairs follow fontconfig's
// 30-metric-aliases.conf; the remaining look-alikes were chosen with
// LibreOffice's VCL.xcu font replacement lists as a guide, restricted to
// families the Google-compatible CSS services carry.
export const WORD_FAMILY_SUBSTITUTES: Record<string, WordFamilySubstitute> = {
  arial: sans('Arimo'),
  helvetica: sans('Arimo'),
  'helvetica neue': sans('Arimo'),
  'arial narrow': sans('Archivo Narrow'),
  'arial black': sans('Archivo Black'),
  calibri: sans('Carlito'),
  'calibri light': sans('Carlito'),
  cambria: serif('Caladea'),
  'times new roman': serif('Tinos'),
  times: serif('Tinos'),
  'courier new': mono('Cousine'),
  courier: mono('Cousine'),
  georgia: serif('Gelasio'),
  'comic sans ms': { remote: 'Comic Neue', generic: 'cursive' },
  verdana: sans('Open Sans'),
  tahoma: sans('Open Sans'),
  'segoe ui': sans('Open Sans'),
  'lucida sans': sans('Open Sans'),
  'lucida sans unicode': sans('Open Sans'),
  'lucida grande': sans('Open Sans'),
  'trebuchet ms': sans('Fira Sans'),
  'franklin gothic': sans('Libre Franklin'),
  'franklin gothic book': sans('Libre Franklin'),
  'franklin gothic medium': sans('Libre Franklin'),
  'century gothic': sans('Questrial'),
  impact: sans('Anton'),
  garamond: serif('EB Garamond'),
  'palatino linotype': serif('EB Garamond'),
  palatino: serif('EB Garamond'),
  'book antiqua': serif('EB Garamond'),
  consolas: mono('Inconsolata'),
  'lucida console': mono('Inconsolata'),
  monaco: mono('Fira Mono'),

  'ms mincho': serif('Noto Serif JP'),
  'ms pmincho': serif('Noto Serif JP'),
  'ｍｓ 明朝': serif('Noto Serif JP'),
  'ｍｓ ｐ明朝': serif('Noto Serif JP'),
  'yu mincho': serif('Noto Serif JP'),
  游明朝: serif('Noto Serif JP'),
  'ms gothic': sans('Noto Sans JP'),
  'ms pgothic': sans('Noto Sans JP'),
  'ms ui gothic': sans('Noto Sans JP'),
  'ｍｓ ゴシック': sans('Noto Sans JP'),
  'ｍｓ ｐゴシック': sans('Noto Sans JP'),
  meiryo: sans('Noto Sans JP'),
  メイリオ: sans('Noto Sans JP'),
  'yu gothic': sans('Noto Sans JP'),
  游ゴシック: sans('Noto Sans JP'),

  simsun: serif('Noto Serif SC'),
  nsimsun: serif('Noto Serif SC'),
  宋体: serif('Noto Serif SC'),
  fangsong: serif('Noto Serif SC'),
  仿宋: serif('Noto Serif SC'),
  kaiti: serif('Noto Serif SC'),
  楷体: serif('Noto Serif SC'),
  simhei: sans('Noto Sans SC'),
  黑体: sans('Noto Sans SC'),
  'microsoft yahei': sans('Noto Sans SC'),
  微软雅黑: sans('Noto Sans SC'),
  dengxian: sans('Noto Sans SC'),
  等线: sans('Noto Sans SC'),

  pmingliu: serif('Noto Serif TC'),
  mingliu: serif('Noto Serif TC'),
  新細明體: serif('Noto Serif TC'),
  細明體: serif('Noto Serif TC'),
  'dfkai-sb': serif('Noto Serif TC'),
  標楷體: serif('Noto Serif TC'),
  'microsoft jhenghei': sans('Noto Sans TC'),
  微軟正黑體: sans('Noto Sans TC'),

  'malgun gothic': sans('Noto Sans KR'),
  '맑은 고딕': sans('Noto Sans KR'),
  gulim: sans('Noto Sans KR'),
  굴림: sans('Noto Sans KR'),
  dotum: sans('Noto Sans KR'),
  돋움: sans('Noto Sans KR'),
  batang: serif('Noto Serif KR'),
  바탕: serif('Noto Serif KR'),
  gungsuh: serif('Noto Serif KR'),
  궁서: serif('Noto Serif KR'),
};

/** The table entry for an Office family name (case-insensitive), or undefined. */
export function resolveWordFamilySubstitute(family: string): WordFamilySubstitute | undefined {
  const key = family.trim().toLowerCase();
  return Object.prototype.hasOwnProperty.call(WORD_FAMILY_SUBSTITUTES, key) ? WORD_FAMILY_SUBSTITUTES[key] : undefined;
}
