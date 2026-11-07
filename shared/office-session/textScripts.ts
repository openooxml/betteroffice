const INTL: ReadonlyArray<readonly [string, ReadonlyArray<readonly [number, number]>]> = [
  ['thai', [[0x0e00, 0x0e7f]]],
  ['lao', [[0x0e80, 0x0eff]]],
  ['khmer', [[0x1780, 0x17ff], [0x19e0, 0x19ff]]],
  ['myanmar', [[0x1000, 0x109f], [0xa9e0, 0xa9ff], [0xaa60, 0xaa7f], [0x116d0, 0x116ff]]],
  ['devanagari', [[0x0900, 0x097f], [0xa8e0, 0xa8ff], [0x11b00, 0x11b5f]]],
  ['bengali', [[0x0980, 0x09ff]]],
  ['gurmukhi', [[0x0a00, 0x0a7f]]],
  ['gujarati', [[0x0a80, 0x0aff]]],
  ['oriya', [[0x0b00, 0x0b7f]]],
  ['tamil', [[0x0b80, 0x0bff], [0x11fc0, 0x11fff]]],
  ['telugu', [[0x0c00, 0x0c7f]]],
  ['kannada', [[0x0c80, 0x0cff]]],
  ['malayalam', [[0x0d00, 0x0d7f]]],
  ['sinhala', [[0x0d80, 0x0dff], [0x111e0, 0x111ff]]],
  ['armenian', [[0x0530, 0x058f], [0xfb13, 0xfb17]]],
  ['georgian', [[0x10a0, 0x10ff], [0x1c90, 0x1cbf], [0x2d00, 0x2d2f]]],
  ['ethiopic', [[0x1200, 0x137f], [0x1380, 0x139f], [0x2d80, 0x2ddf], [0xab00, 0xab2f], [0x1e7e0, 0x1e7ff]]],
  ['tibetan', [[0x0f00, 0x0fff]]],
  ['mongolian', [[0x1800, 0x18af], [0x11660, 0x1167f]]],
  ['syriac', [[0x0700, 0x074f], [0x0860, 0x086f]]],
  ['thaana', [[0x0780, 0x07bf]]],
  ['cherokee', [[0x13a0, 0x13ff], [0xab70, 0xabbf]]],
  ['tifinagh', [[0x2d30, 0x2d7f]]],
];

const HEBREW: ReadonlyArray<readonly [number, number]> = [[0x0590, 0x05ff], [0xfb1d, 0xfb4f]];
const ARABIC: ReadonlyArray<readonly [number, number]> = [[0x0600, 0x06ff], [0x0750, 0x077f], [0x0870, 0x08ff], [0xfb50, 0xfdff], [0xfe70, 0xfeff]];
const HANGUL: ReadonlyArray<readonly [number, number]> = [[0x1100, 0x11ff], [0x3130, 0x318f], [0xa960, 0xa97f], [0xac00, 0xd7ff], [0xffa0, 0xffdc]];
const KANA: ReadonlyArray<readonly [number, number]> = [[0x3040, 0x30ff], [0x31f0, 0x31ff], [0xff66, 0xff9f]];
const HAN: ReadonlyArray<readonly [number, number]> = [
  [0x3000, 0x303f], [0x3400, 0x4dbf], [0x4e00, 0x9fff], [0xf900, 0xfaff], [0xfe30, 0xfe4f], [0xff00, 0xff65], [0x20000, 0x3ffff],
];

const within = (ranges: ReadonlyArray<readonly [number, number]>, point: number) => ranges.some(([low, high]) => point >= low && point <= high);

/**
 * The script coverage buckets of one text, in the order the native engines ask their font manager for faces
 * (`ooxml_text::paragraph_scripts_with_han_fallback(&[text], false)`): Japanese, Korean, Simplified Chinese for
 * Han without kana or Hangul, Arabic, Hebrew, then other scripts as they first appear.
 */
export function textScripts(text: string): string[] {
  let han = false;
  let kana = false;
  let hangul = false;
  let arabic = false;
  let hebrew = false;
  const intl: string[] = [];
  for (const character of text) {
    const point = character.codePointAt(0)!;
    if (within(HEBREW, point)) hebrew = true;
    else if (within(ARABIC, point)) arabic = true;
    else if (within(HANGUL, point)) hangul = true;
    else if (within(KANA, point)) kana = true;
    else if (within(HAN, point)) han = true;
    const script = INTL.find(([, ranges]) => within(ranges, point))?.[0];
    if (script && !intl.includes(script)) intl.push(script);
  }
  const scripts: string[] = [];
  if (kana) scripts.push('cjk-jp');
  if (hangul) scripts.push('cjk-kr');
  if (han && !kana && !hangul) scripts.push('cjk-sc');
  if (arabic) scripts.push('arabic');
  if (hebrew) scripts.push('hebrew');
  scripts.push(...intl);
  return scripts;
}
