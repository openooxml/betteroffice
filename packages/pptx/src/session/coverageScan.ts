const MAX_PART_BYTES = 4 * 1024 * 1024;

interface ZipEntry {
  method: number;
  compressedSize: number;
  size: number;
  offset: number;
}

function centralDirectory(bytes: Uint8Array): Map<string, ZipEntry> | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const floor = Math.max(0, bytes.byteLength - 22 - 0xffff);
  let end = -1;
  for (let at = bytes.byteLength - 22; at >= floor; at--) {
    if (view.getUint32(at, true) === 0x06054b50) { end = at; break; }
  }
  if (end < 0) return undefined;
  const count = view.getUint16(end + 10, true);
  let at = view.getUint32(end + 16, true);
  if (count === 0xffff || at === 0xffffffff) return undefined;
  const decoder = new TextDecoder();
  const entries = new Map<string, ZipEntry>();
  for (let index = 0; index < count; index++) {
    if (at + 46 > bytes.byteLength || view.getUint32(at, true) !== 0x02014b50) return undefined;
    const nameLength = view.getUint16(at + 28, true);
    const skip = nameLength + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
    const name = decoder.decode(bytes.subarray(at + 46, at + 46 + nameLength));
    entries.set(name, {
      method: view.getUint16(at + 10, true),
      compressedSize: view.getUint32(at + 20, true),
      size: view.getUint32(at + 24, true),
      offset: view.getUint32(at + 42, true),
    });
    at += 46 + skip;
  }
  return entries;
}

async function readPart(bytes: Uint8Array, entries: Map<string, ZipEntry>, name: string): Promise<string | undefined> {
  const entry = entries.get(name);
  if (!entry || entry.size > MAX_PART_BYTES) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (entry.offset + 30 > bytes.byteLength || view.getUint32(entry.offset, true) !== 0x04034b50) return undefined;
  const start = entry.offset + 30 + view.getUint16(entry.offset + 26, true) + view.getUint16(entry.offset + 28, true);
  const data = bytes.slice(start, start + entry.compressedSize);
  if (entry.method === 0) return new TextDecoder().decode(data);
  if (entry.method !== 8 || typeof DecompressionStream === 'undefined') return undefined;
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new TextDecoder().decode(await new Response(stream).arrayBuffer());
}

function relationshipTarget(rels: string, id: string): string | undefined {
  for (const match of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    if (new RegExp(`\\bId="${id}"`).test(match[0])) return /\bTarget="([^"]+)"/.exec(match[0])?.[1];
  }
  return undefined;
}

function hanScript(language: string | undefined): string | undefined {
  if (!language) return undefined;
  const [primary, ...rest] = language.toLowerCase().split(/[-_]/);
  if (primary === 'ja') return 'cjk-jp';
  if (primary === 'ko') return 'cjk-kr';
  if (primary !== 'zh') return undefined;
  if (rest.includes('hans')) return 'cjk-sc';
  return rest.some((part) => part === 'hant' || part === 'tw' || part === 'hk' || part === 'mo') ? 'cjk-tc' : 'cjk-sc';
}

/** CJK coverage scripts the slide's text certainly needs: kana, hangul, or Han in a run tagged ja, ko or zh. */
export function slideCoverageScripts(xml: string): string[] {
  const scripts = new Set<string>();
  let language: string | undefined;
  for (const match of xml.matchAll(/<a:(?:rPr|endParaRPr)\b([^>]*)>|<a:t>([^<]*)<\/a:t>|<\/a:r>/g)) {
    if (match[1] !== undefined) { language = /\blang="([^"]*)"/.exec(match[1])?.[1]; continue; }
    if (match[2] === undefined) { language = undefined; continue; }
    for (const character of match[2]) {
      const point = character.codePointAt(0)!;
      if ((point >= 0x3040 && point <= 0x30ff) || (point >= 0x31f0 && point <= 0x31ff) || (point >= 0xff66 && point <= 0xff9f)) {
        scripts.add('cjk-jp');
      } else if ((point >= 0xac00 && point <= 0xd7ff) || (point >= 0x1100 && point <= 0x11ff) || (point >= 0x3130 && point <= 0x318f)) {
        scripts.add('cjk-kr');
      } else if ((point >= 0x4e00 && point <= 0x9fff) || (point >= 0x3400 && point <= 0x4dbf) || (point >= 0xf900 && point <= 0xfaff) || point >= 0x20000) {
        const script = hanScript(language);
        if (script) scripts.add(script);
      }
    }
  }
  return [...scripts];
}

/** Coverage scripts of the slide that opens first, from the package bytes; empty when anything is unexpected. */
export async function activeSlideCoverageScripts(source: ArrayBuffer | Uint8Array, activeSlideIndex = 0): Promise<string[]> {
  try {
    const bytes = source instanceof Uint8Array ? source : new Uint8Array(source);
    const entries = centralDirectory(bytes);
    if (!entries) return [];
    const presentation = await readPart(bytes, entries, 'ppt/presentation.xml');
    const rels = await readPart(bytes, entries, 'ppt/_rels/presentation.xml.rels');
    if (!presentation || !rels) return [];
    const ids = [...(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/.exec(presentation)?.[1] ?? '').matchAll(/\br:id="([^"]+)"/g)]
      .map((match) => match[1]);
    const id = ids[Math.max(0, Math.min(activeSlideIndex, ids.length - 1))];
    const target = id === undefined ? undefined : relationshipTarget(rels, id);
    if (!target) return [];
    const path = target.startsWith('/') ? target.slice(1) : `ppt/${target.replace(/^\.\//, '')}`;
    const xml = await readPart(bytes, entries, path);
    return xml ? slideCoverageScripts(xml) : [];
  } catch {
    return [];
  }
}
