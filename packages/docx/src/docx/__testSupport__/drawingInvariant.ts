import { unzipContainer } from '../../wasm/opc';

const DRAWING_ELEMENTS = ['w:drawing', 'w:pict', 'w:object', 'mc:AlternateContent'] as const;

/** Per part, how many drawing containers of each kind it holds. */
export function countDrawings(bytes: Uint8Array): Map<string, Record<string, number>> {
  const parts = new Map<string, Record<string, number>>();
  for (const [part, data] of Object.entries(unzipContainer(bytes))) {
    if (!part.startsWith('word/') || !part.endsWith('.xml')) continue;
    const xml = new TextDecoder().decode(data);
    const counts: Record<string, number> = {};
    for (const name of DRAWING_ELEMENTS) {
      const found = xml.match(new RegExp(`<${name}(?=[\\s/>])`, 'g'))?.length ?? 0;
      if (found > 0) counts[name] = found;
    }
    if (Object.keys(counts).length > 0) parts.set(part, counts);
  }
  return parts;
}

/** Throws when the saved package holds fewer drawing containers of any kind in any part than the source. */
export function assertDrawingInvariant(source: Uint8Array, saved: Uint8Array): void {
  const written = countDrawings(saved);
  const missing: string[] = [];
  for (const [part, counts] of countDrawings(source)) {
    for (const [name, count] of Object.entries(counts)) {
      const lost = count - (written.get(part)?.[name] ?? 0);
      if (lost > 0) missing.push(`${part}: ${name} missing ${lost}`);
    }
  }
  if (missing.length > 0) throw new Error(`Lost drawings:\n${missing.join('\n')}`);
}
