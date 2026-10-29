import type { FontDecoder } from '@betteroffice/fonts';
import type { PptxFontFace } from '@betteroffice/pptx';

export async function prepareDefaultFonts(
  runtime: { decoder: FontDecoder }, families: () => Promise<string[]>, report: () => void,
): Promise<PptxFontFace[]> {
  const requirements = families();
  void requirements.catch(() => {});
  let fonts: typeof import('@betteroffice/fonts');
  try {
    fonts = await import('@betteroffice/fonts');
  } catch (error) {
    const message = (error as { message?: unknown } | null)?.message;
    if (typeof message === 'string' && /^(?:Cannot find (?:module|package)|Failed to resolve module specifier|Error resolving module specifier|The specifier|Module name,) ["'“]@betteroffice\/fonts["'”]/.test(message)) return [];
    throw error;
  }
  const provider = fonts.createFontProvider(undefined, runtime);
  const pending = new Map<string, Promise<PptxFontFace>>();
  const load = (family: string) => {
    for (const bold of [false, true]) for (const italic of [false, true]) {
      const key = `${family}|${bold}|${italic}`;
      if (pending.has(key)) continue;
      const loader = provider.resolve(family, bold, italic) ?? provider.resolveFamily(family, bold, italic);
      if (loader) {
        const face = loader().then(bytes => ({ family, bold, italic, bytes: new Uint8Array(bytes) }));
        void face.catch(() => {});
        pending.set(key, face);
      }
    }
  };
  load('Arial');
  for (const family of await requirements) load(family);
  const results = await Promise.allSettled(pending.values());
  if (results.some(result => result.status === 'rejected')) report();
  return results.flatMap(result => result.status === 'fulfilled' ? [result.value] : []);
}
