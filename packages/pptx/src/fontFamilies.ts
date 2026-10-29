export function presentationFontFamilies(document: unknown): string[] {
  const families = new Set<string>(['Arial']);
  function visit(value: unknown): void {
    if (value === null || typeof value !== 'object') return;
    for (const [key, child] of Object.entries(value)) {
      if (['fontFamily', 'latin', 'ea', 'cs'].includes(key) && typeof child === 'string' && child && !child.startsWith('+')) {
        families.add(child);
      } else visit(child);
    }
  }
  visit(document);
  return [...families];
}
