import { unzipContainer } from '../../wasm/opc';

type Counts = { Start: number; End: number };
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function decodedId(value: string): string {
  const decoded = value.replace(/&#(?:x([0-9a-f]+)|([0-9]+));|&(amp|lt|gt|quot|apos);/gi, (entity, hex, decimal, named) => {
    if (named) return ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" } as Record<string, string>)[named] ?? entity;
    const code = Number.parseInt(hex ?? decimal, hex === undefined ? 10 : 16);
    return code <= 0x10ffff ? String.fromCodePoint(code) : entity;
  });
  return /^[-+]?\d+$/.test(decoded) && Number.isSafeInteger(Number(decoded)) ? String(Number(decoded)) : decoded;
}

export function collectBookmarkBoundaries(bytes: Uint8Array): Map<string, Map<string, Counts>> {
  const parts = new Map<string, Map<string, Counts>>();
  for (const [part, partBytes] of Object.entries(unzipContainer(bytes))) {
    if (!part.endsWith('.xml')) continue;
    const xml = new TextDecoder().decode(partBytes);
    const ids = new Map<string, Counts>();
    const scopes: Array<Map<string, string>> = [];
    for (const [tag] of xml.matchAll(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<[^<>"']*(?:"[^"]*"[^<>"']*|'[^']*'[^<>"']*)*>/g)) {
      if (tag.startsWith('<!') || tag.startsWith('<?')) continue;
      if (tag.startsWith('</')) { scopes.pop(); continue; }
      const name = /^<([^\s/>]+)/.exec(tag)?.[1];
      if (!name) continue;
      const scope = new Map(scopes.at(-1));
      const attributes = [...tag.matchAll(/([^\s=<>/]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)]
        .map(attribute => [attribute[1]!, decodedId(attribute[2] ?? attribute[3]!)] as const);
      for (const [key, value] of attributes) {
        if (key === 'xmlns') scope.set('', value);
        else if (key.startsWith('xmlns:')) scope.set(key.slice(6), value);
      }
      if (!/\/\s*>$/.test(tag)) scopes.push(scope);
      const [prefix, local] = name.includes(':') ? name.split(':') : ['', name];
      const kind = /^bookmark(Start|End)$/.exec(local!)?.[1] as keyof Counts | undefined;
      if (!kind || scope.get(prefix!) !== W) continue;
      for (const [key, id] of attributes) {
        const [prefix, local] = key.split(':');
        if (local !== 'id' || scope.get(prefix!) !== W) continue;
        const counts = ids.get(id) ?? { Start: 0, End: 0 };
        counts[kind] += 1;
        ids.set(id, counts);
        break;
      }
    }
    if (ids.size > 0) parts.set(part, ids);
  }
  return parts;
}

export function assertBookmarkInvariant(source: Uint8Array, saved: Uint8Array): void {
  const written = collectBookmarkBoundaries(saved);
  const missing: string[] = [];
  for (const [part, ids] of collectBookmarkBoundaries(source)) {
    for (const [id, counts] of ids) {
      for (const kind of ['Start', 'End'] as const) {
        const lost = counts[kind] - (written.get(part)?.get(id)?.[kind] ?? 0);
        if (lost > 0) missing.push(`${part}: w:id=${id} missing ${kind} (${lost})`);
      }
    }
  }
  if (missing.length > 0) throw new Error(`Lost bookmark boundaries:\n${missing.join('\n')}`);
}
