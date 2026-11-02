const MAX_PART_BYTES = 4 * 1024 * 1024;
const MAX_SCAN_BYTES = 8 * 1024 * 1024;
const MAX_ENTRIES = 5000;
const INFLATE_INPUT_BYTES = 1024;

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
  if (count > MAX_ENTRIES || at === 0xffffffff) return undefined;
  const decoder = new TextDecoder();
  const entries = new Map<string, ZipEntry>();
  for (let index = 0; index < count; index++) {
    if (at + 46 > bytes.byteLength || view.getUint32(at, true) !== 0x02014b50) return undefined;
    const nameLength = view.getUint16(at + 28, true);
    const skip = nameLength + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
    if (at + 46 + skip > bytes.byteLength) return undefined;
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

/**
 * Inflates `data`, giving up as soon as the output passes `limit` bytes, whatever the entry declared. Input goes
 * in small slices under backpressure, so one slice expands to at most about a megabyte before the count sees it.
 */
async function inflate(data: Uint8Array<ArrayBuffer>, limit: number): Promise<Uint8Array | undefined> {
  const stream = new DecompressionStream('deflate-raw');
  const writer = stream.writable.getWriter();
  const reader = stream.readable.getReader();
  let stopped = false;
  void (async () => {
    for (let at = 0; at < data.byteLength && !stopped; at += INFLATE_INPUT_BYTES) {
      await writer.ready;
      if (!stopped) await writer.write(data.subarray(at, at + INFLATE_INPUT_BYTES));
    }
    if (!stopped) await writer.close();
  })().catch(() => {});
  const chunks: Uint8Array[] = [];
  let total = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > limit) {
        stopped = true;
        void reader.cancel().catch(() => {});
        void writer.abort().catch(() => {});
        return undefined;
      }
      chunks.push(value);
    }
  } catch {
    stopped = true;
    void writer.abort().catch(() => {});
    return undefined;
  }
  const out = new Uint8Array(total);
  let at = 0;
  for (const chunk of chunks) { out.set(chunk, at); at += chunk.byteLength; }
  return out;
}

async function readPart(
  bytes: Uint8Array, entries: Map<string, ZipEntry>, name: string, budget: { left: number },
): Promise<string | undefined> {
  const entry = entries.get(name);
  const limit = Math.min(MAX_PART_BYTES, budget.left);
  if (!entry || entry.size > limit || entry.compressedSize > limit) return undefined;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (entry.offset + 30 > bytes.byteLength || view.getUint32(entry.offset, true) !== 0x04034b50) return undefined;
  const start = entry.offset + 30 + view.getUint16(entry.offset + 26, true) + view.getUint16(entry.offset + 28, true);
  if (start + entry.compressedSize > bytes.byteLength) return undefined;
  const data = bytes.slice(start, start + entry.compressedSize);
  let out: Uint8Array | undefined;
  if (entry.method === 0) out = entry.compressedSize === entry.size ? data : undefined;
  else if (entry.method === 8 && typeof DecompressionStream !== 'undefined') out = await inflate(data, limit);
  if (!out) return undefined;
  budget.left -= out.byteLength;
  return new TextDecoder().decode(out);
}

/** Calls `visit` with each tag's text between `<` and `>` in one pass; stops at an unterminated tag. */
function eachTag(xml: string, from: number, to: number, visit: (tag: string, end: number) => number | void): void {
  let at = from;
  while (at < to) {
    const open = xml.indexOf('<', at);
    if (open < 0 || open >= to) return;
    const close = xml.indexOf('>', open + 1);
    if (close < 0) return;
    const next = visit(xml.slice(open + 1, close), close + 1);
    at = typeof next === 'number' ? next : close + 1;
  }
}

const isSpace = (code: number) => code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d;

function isTag(tag: string, name: string): boolean {
  if (!tag.startsWith(name)) return false;
  const after = tag.charCodeAt(name.length);
  return Number.isNaN(after) || after === 0x2f || isSpace(after);
}

/** The value of attribute `name` in a tag's text, read as a plain string. */
function attribute(tag: string, name: string): string | undefined {
  for (let at = tag.indexOf(`${name}=`); at >= 0; at = tag.indexOf(`${name}=`, at + 1)) {
    const quote = tag[at + name.length + 1];
    if (!isSpace(tag.charCodeAt(at - 1)) || (quote !== '"' && quote !== "'")) continue;
    const start = at + name.length + 2;
    const end = tag.indexOf(quote, start);
    return end < 0 ? undefined : tag.slice(start, end);
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

function addScripts(text: string, han: string | undefined, scripts: Set<string>): void {
  for (const character of text) {
    const point = character.codePointAt(0)!;
    if ((point >= 0x3040 && point <= 0x30ff) || (point >= 0x31f0 && point <= 0x31ff) || (point >= 0xff66 && point <= 0xff9f)) {
      scripts.add('cjk-jp');
    } else if ((point >= 0xac00 && point <= 0xd7ff) || (point >= 0x1100 && point <= 0x11ff) || (point >= 0x3130 && point <= 0x318f)) {
      scripts.add('cjk-kr');
    } else if ((point >= 0x4e00 && point <= 0x9fff) || (point >= 0x3400 && point <= 0x4dbf) || (point >= 0xf900 && point <= 0xfaff) || point >= 0x20000) {
      if (han) scripts.add(han);
    }
  }
}

/** CJK coverage scripts the slide's text certainly needs: kana, hangul, or Han in a run tagged ja, ko or zh. */
export function slideCoverageScripts(xml: string): string[] {
  const scripts = new Set<string>();
  let han: string | undefined;
  eachTag(xml, 0, xml.length, (tag, end) => {
    if (isTag(tag, 'a:rPr') || isTag(tag, 'a:endParaRPr')) {
      han = hanScript(attribute(tag, 'lang'));
    } else if (tag === '/a:r') {
      han = undefined;
    } else if (isTag(tag, 'a:t') && !tag.endsWith('/')) {
      const close = xml.indexOf('<', end);
      const stop = close < 0 ? xml.length : close;
      addScripts(xml.slice(end, stop), han, scripts);
      return stop;
    }
  });
  return [...scripts];
}

function slideIds(presentation: string): string[] {
  const ids: string[] = [];
  const start = presentation.indexOf('<p:sldIdLst');
  if (start < 0) return ids;
  const end = presentation.indexOf('</p:sldIdLst>', start);
  eachTag(presentation, start, end < 0 ? presentation.length : end, (tag) => {
    if (!isTag(tag, 'p:sldId')) return;
    const id = attribute(tag, 'r:id');
    if (id !== undefined) ids.push(id);
  });
  return ids;
}

function relationshipTarget(rels: string, id: string): string | undefined {
  let target: string | undefined;
  eachTag(rels, 0, rels.length, (tag) => {
    if (!isTag(tag, 'Relationship') || attribute(tag, 'Id') !== id) return;
    target = attribute(tag, 'Target');
    return rels.length;
  });
  return target;
}

export interface OpeningSlideScan {
  scripts: string[];
  openingSlideMetafile: boolean;
  packageMetafile: boolean;
}

function metafileTarget(target: string): boolean {
  const lower = target.toLowerCase();
  return lower.endsWith('.emf') || lower.endsWith('.wmf');
}

export async function scanOpeningSlide(source: ArrayBuffer | Uint8Array, activeSlideIndex = 0): Promise<OpeningSlideScan> {
  const result: OpeningSlideScan = { scripts: [], openingSlideMetafile: false, packageMetafile: false };
  try {
    const bytes = source instanceof Uint8Array ? source : new Uint8Array(source);
    const entries = centralDirectory(bytes);
    if (!entries) return result;
    result.packageMetafile = [...entries.keys()].some((name) => name.startsWith('ppt/media/') && metafileTarget(name));
    const budget = { left: MAX_SCAN_BYTES };
    const presentation = await readPart(bytes, entries, 'ppt/presentation.xml', budget);
    const rels = await readPart(bytes, entries, 'ppt/_rels/presentation.xml.rels', budget);
    if (!presentation || !rels) return result;
    const ids = slideIds(presentation);
    const id = ids[Math.max(0, Math.min(activeSlideIndex, ids.length - 1))];
    const target = id === undefined ? undefined : relationshipTarget(rels, id);
    if (!target) return result;
    const relative = target.startsWith('./') ? target.slice(2) : target;
    const path = target.startsWith('/') ? target.slice(1) : `ppt/${relative}`;
    const slash = path.lastIndexOf('/');
    const slideRels = await readPart(bytes, entries, `${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`, budget);
    if (slideRels) eachTag(slideRels, 0, slideRels.length, (tag) => {
      if (isTag(tag, 'Relationship') && metafileTarget(attribute(tag, 'Target') ?? '')) result.openingSlideMetafile = true;
    });
    const xml = await readPart(bytes, entries, path, budget);
    if (xml) result.scripts = slideCoverageScripts(xml);
    return result;
  } catch {
    return { scripts: [], openingSlideMetafile: false, packageMetafile: false };
  }
}

export async function activeSlideMetafileFlags(source: ArrayBuffer | Uint8Array, activeSlideIndex = 0): Promise<{
  openingSlideMetafile: boolean; packageMetafile: boolean;
}> {
  const { openingSlideMetafile, packageMetafile } = await scanOpeningSlide(source, activeSlideIndex);
  return { openingSlideMetafile, packageMetafile };
}

/** Coverage scripts of the slide that opens first. */
export async function activeSlideCoverageScripts(source: ArrayBuffer | Uint8Array, activeSlideIndex = 0): Promise<string[]> {
  return (await scanOpeningSlide(source, activeSlideIndex)).scripts;
}
