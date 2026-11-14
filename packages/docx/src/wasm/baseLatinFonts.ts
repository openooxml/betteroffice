const installed = new WeakMap<FontFaceSet, Map<string, Promise<void>>>();

export function installBaseLatinFace(bytes: ArrayBuffer, bold: boolean, italic: boolean): Promise<void> {
  const fonts = typeof document === 'undefined'
    ? (globalThis as { fonts?: FontFaceSet }).fonts : document.fonts;
  if (!fonts || typeof FontFace !== 'function') return Promise.resolve();
  let faces = installed.get(fonts);
  if (!faces) { faces = new Map(); installed.set(fonts, faces); }
  const key = `${bold}|${italic}`;
  const existing = faces.get(key);
  if (existing) return existing;
  const pending = Promise.resolve().then(async () => {
    const face = new FontFace('BetterOffice Base Latin', bytes.slice(0), {
      weight: bold ? '700' : '400', style: italic ? 'italic' : 'normal',
    });
    await face.load();
    fonts.add(face);
  }).catch(() => { faces!.delete(key); });
  faces.set(key, pending);
  return pending;
}
