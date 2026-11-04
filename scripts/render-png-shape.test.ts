import { describe, expect, test } from 'bun:test';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const ROOT = resolve(import.meta.dir, '..');
const FORMATS = ['docx', 'pptx', 'vsdx', 'xlsx'] as const;
type Format = (typeof FORMATS)[number];

const JS_TARGET: Record<Format, RegExp> = {
  docx: /^\{\s*page: number;?\s*\}$/,
  pptx: /^\{\s*slide: number(?: \| string)?;?\s*\}$/,
  vsdx: /^\{\s*page: number;?\s*\}$/,
  xlsx: /^\|?\s*\{\s*sheet: number(?: \| string)?;\s*viewport: \w+;?\s*\}\s*\|\s*\{\s*sheet: number(?: \| string)?;\s*range: (?!number\b|boolean\b)\w+(?: \| '\w+')?;?\s*\}$/,
};

function aliasBody(text: string, name: string): string | undefined {
  const start = text.search(new RegExp(`\\btype ${name}\\s*=`));
  if (start < 0) return undefined;
  const from = text.indexOf('=', start) + 1;
  let depth = 0;
  for (let at = from; at < text.length; at++) {
    if (text[at] === '{') depth++;
    else if (text[at] === '}') depth--;
    else if (text[at] === ';' && depth === 0) return text.slice(from, at).trim();
  }
  return undefined;
}

export function jsMatches(format: Format, text: string): boolean {
  const declared = text.match(/renderPng\(\s*target: ([^,]+?),\s*options\?: [^)]+\): Promise<Uint8Array>\s*[;{]/)?.[1].trim();
  if (!declared) return false;
  const body = /^\w+$/.test(declared) ? aliasBody(text, declared) : declared;
  return body !== undefined && JS_TARGET[format].test(body);
}

export function pythonMatches(text: string): boolean {
  return /def render_png\(\s*self,\s*target(?::(?!\s*(?:int|str|float)\b)[^,)]+)?,\s*\*,[^)]*\bscale\b[^)]*\)\s*->\s*(?:bytes|Png)\s*:/.test(text);
}

export function rustMatches(text: string): boolean {
  return /pub fn render_png\(\s*&self,\s*target: (?!&?(?:[ui](?:8|16|32|64|128|size)|f32|f64|str|String)\b)[^,]+,\s*options: [^,)]+,?\s*\)\s*->\s*(?:\w+::)*Result<(?:Vec<u8>|Rendered\w+)>\s*\{/.test(text);
}

function sources(root: string): string[] {
  const directory = resolve(ROOT, root);
  if (!existsSync(directory)) return [];
  return [...new Bun.Glob('**/*.ts').scanSync(directory)]
    .filter((path) => !/\.test\.|\.d\.ts$|generated\//.test(path))
    .map((path) => readFileSync(resolve(directory, path), 'utf8'));
}

function jsShape(format: Format): boolean {
  return jsMatches(format, sources(`packages/${format}/src`).join('\n'));
}

function pythonShape(format: Format): boolean {
  const path = resolve(ROOT, `bindings/python-${format}/python/betteroffice_${format}/_betteroffice_${format}.pyi`);
  return existsSync(path) && pythonMatches(readFileSync(path, 'utf8'));
}

function rustShape(format: Format): boolean {
  const root = resolve(ROOT, `crates/betteroffice-${format}/src`);
  return existsSync(root) && [...new Bun.Glob('**/*.rs').scanSync(root)]
    .some((path) => rustMatches(readFileSync(resolve(root, path), 'utf8')));
}

function nativeRenderPng(format: Format): boolean {
  const path = resolve(ROOT, `bindings/${format}-native/index.d.ts`);
  return existsSync(resolve(ROOT, `bindings/${format}-native/package.json`)) && existsSync(path) &&
    /\brenderPng(?:Target)?\(/.test(readFileSync(path, 'utf8'));
}

const CHECKS: Record<string, () => boolean> = Object.fromEntries(FORMATS.flatMap((format) => [
  [`${format} js`, () => jsShape(format)],
  [`${format} native renderPng`, () => nativeRenderPng(format)],
  [`${format} python`, () => pythonShape(format)],
  [`${format} rust`, () => rustShape(format)],
]));

const PENDING = [
  'docx js',
  'docx native renderPng',
  'vsdx native renderPng',
];

describe('the shape checks tell right from wrong', () => {
  const js = 'renderPng(target: { page: number }, options?: RenderPngOptions): Promise<Uint8Array>;';
  test.each([
    ['docx', js, true],
    ['vsdx', js, true],
    ['pptx', js, false],
    ['pptx', 'renderPng(target: { slide: number }, options?: RenderPngOptions): Promise<Uint8Array>;', true],
    ['pptx', 'type T = { slide: number | string };\nrenderPng(target: T, options?: RenderPngOptions): Promise<Uint8Array>;', true],
    ['pptx', 'renderPng(target: { slide: boolean }, options?: RenderPngOptions): Promise<Uint8Array>;', false],
    ['docx', 'renderPng(index: number, options?: RenderPngOptions): Promise<Uint8Array>;', false],
    ['docx', 'renderPng(target: { page: number }): Promise<Uint8Array>;', false],
    ['docx', 'renderPng(target: { page: number }, options?: RenderPngOptions): Uint8Array;', false],
    ['docx', 'renderPng(target: { page: number }, options?: RenderPngOptions): Promise<Uint8Array> | undefined;', false],
    ['xlsx', 'type Target = { sheet: number; viewport: Viewport } | { sheet: number; range: string };\nrenderPng(target: Target, options?: O): Promise<Uint8Array>;', true],
    ['xlsx', 'type Target = {\n  sheet: number;\n  viewport: Viewport;\n} | {\n  sheet: number;\n  range: string;\n};\nrenderPng(target: Target, options?: O): Promise<Uint8Array>;', true],
    ['xlsx', 'renderPng(target: { sheet: number; viewport: Viewport } | { sheet: number; range: string }, options?: O): Promise<Uint8Array>;', true],
    ['xlsx', 'type Target = { sheet: number; viewport: Viewport };\nrenderPng(target: Target, options?: O): Promise<Uint8Array>;', false],
    ['xlsx', 'renderPng(target: Missing, options?: O): Promise<Uint8Array>;', false],
    ['xlsx', 'type Target =\n  | { sheet: number | string; viewport: Viewport }\n  | { sheet: number | string; range: XlsxRangeAddress };\nrenderPng(target: Target, options?: O): Promise<Uint8Array>;', true],
    ['xlsx', 'type Target =\n  | { sheet: number | string; viewport: Viewport }\n  | { sheet: number | string; range: XlsxRangeAddress | \'used\' };\nrenderPng(target: Target, options?: O): Promise<Uint8Array>;', true],
    ['xlsx', 'type Target = { sheet: number | string; viewport: Viewport } | { sheet: number | string; range: number };\nrenderPng(target: Target, options?: O): Promise<Uint8Array>;', false],
    ['xlsx', 'type Target = { sheet: boolean; viewport: Viewport } | { sheet: boolean; range: string };\nrenderPng(target: Target, options?: O): Promise<Uint8Array>;', false],
  ] as const)('js %s %s', (format, text, expected) => expect(jsMatches(format, text)).toBe(expected));

  test.each([
    ['def render_png(self, target: PageTarget, *, scale: float = 1.0) -> bytes: ...', true],
    ['def render_png(self, target: PngTarget, *, scale: float = ...) -> Png: ...', true],
    ['def render_png(self, target: PngTarget, *, scale: float = ...) -> Png | None: ...', false],
    ['def render_png(\n    self,\n    target,\n    *,\n    scale: float = 1.0,\n    background: str | None = None,\n) -> bytes: ...', true],
    ['def render_png(self, target: int, *, scale: float = 1.0) -> bytes: ...', false],
    ['def render_png(self, target: PageTarget) -> None: ...', false],
    ['def render_png(self, target: PageTarget, *, scale: float = 1.0) -> bytes | None: ...', false],
    ['def render_png(self, display_list: DisplayList, page: int = 0) -> bytes: ...', false],
  ] as const)('python %s', (text, expected) => expect(pythonMatches(text)).toBe(expected));

  test.each([
    ['pub fn render_png(&self, target: PngTarget, options: &RenderOptions) -> Result<Vec<u8>> {', true],
    ['pub fn render_png(\n    &self,\n    target: &SheetTarget,\n    options: &RenderOptions,\n) -> crate::Result<RenderedPng> {', true],
    ['pub fn render_png(&self, target: usize) {', false],
    ['pub fn render_png(&self, target: PngTarget, options: &RenderOptions) -> Result<Option<RenderedPng>> {', false],
    ['pub fn render_png(&self, target: PngTarget, options: &RenderOptions, page: usize) -> Result<RenderedPng> {', false],
    ['pub fn render_png(&self, target: u64, options: &RenderOptions) -> Result<RenderedPng> {', false],
    ['pub fn render_png(&self, target: &str, options: &RenderOptions) -> Result<RenderedPng> {', false],
    ['pub fn render_png(&self, target: usize, options: &RenderOptions) -> Result<Vec<u8>> {', false],
    ['pub fn render_png(&self, viewport: &Viewport) -> Result<RenderedPng> {', false],
  ] as const)('rust %s', (text, expected) => expect(rustMatches(text)).toBe(expected));
});

describe('renderPng has one shape on every target', () => {
  test('pending entries name real targets', () => {
    for (const name of PENDING) expect(Object.keys(CHECKS)).toContain(name);
  });

  for (const [name, check] of Object.entries(CHECKS)) {
    test(name, () => {
      if (PENDING.includes(name)) expect(check(), `${name} now has the shared shape: delete it from PENDING`).toBe(false);
      else expect(check()).toBe(true);
    });
  }
});
