type CellStyle = { builtinId?: number; name?: string; xfId?: number };

export function normalFontIndex(styles: CellStyle[], styleFontIds: number[]) {
  const normal =
    styles.find((style) => style.builtinId === 0) ??
    styles.find((style) => style.name === 'Normal');
  return styleFontIds[normal?.xfId ?? 0] ?? 0;
}
