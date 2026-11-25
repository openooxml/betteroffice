import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

export const GLIBC_FLOOR = '2.17';

const PT_LOAD = 1;
const PT_DYNAMIC = 2;
const DT_NULL = 0n;
const DT_HASH = 4n;
const DT_STRTAB = 5n;
const DT_SYMTAB = 6n;
const DT_STRSZ = 10n;
const DT_SYMENT = 11n;
const DT_GNU_HASH = 0x6ffffef5n;
const DT_VERSYM = 0x6ffffff0n;
const DT_VERNEED = 0x6ffffffen;
const DT_VERNEEDNUM = 0x6fffffffn;
const SHN_UNDEF = 0;
const STB_WEAK = 2;
const SYM_SIZE = 24;
const VERNEED_SIZE = 16;
const VERNAUX_SIZE = 16;

/** @param {string} a @param {string} b */
export function compareGlibc(a, b) {
  const left = a.replace(/^GLIBC_/, '').split('.').map(Number);
  const right = b.replace(/^GLIBC_/, '').split('.').map(Number);
  for (let index = 0; index < Math.max(left.length, right.length); index += 1) {
    const difference = (left[index] ?? 0) - (right[index] ?? 0);
    if (difference) return difference;
  }
  return 0;
}

/** @param {Uint8Array} bytes */
function readDynamic(bytes) {
  if (bytes[0] !== 0x7f || bytes[1] !== 0x45 || bytes[2] !== 0x4c || bytes[3] !== 0x46) {
    throw new Error('expected an ELF file');
  }
  if (bytes[4] !== 2 || bytes[5] !== 1) throw new Error('expected little-endian ELF64');

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (offset) => view.getUint16(offset, true);
  const u32 = (offset) => view.getUint32(offset, true);
  const u64 = (offset) => view.getBigUint64(offset, true);
  const requireRange = (offset, size, label, end = bytes.length) => {
    if (offset > end - size) throw new Error(`truncated ELF ${label}`);
  };

  requireRange(0, 64, 'header');
  const machine = u16(18);
  if (machine !== 62 && machine !== 183) throw new Error('expected x86_64 or aarch64 ELF');
  const headersOffset = Number(u64(32));
  const headerSize = u16(54);
  const headerCount = u16(56);
  if (headerCount === 0) return;
  if (headerSize < 56) throw new Error('invalid ELF program header size');
  requireRange(headersOffset, headerSize * headerCount, 'program headers');

  const loads = [];
  let dynamic;
  for (let index = 0; index < headerCount; index += 1) {
    const header = headersOffset + index * headerSize;
    const type = u32(header);
    if (type !== PT_LOAD && type !== PT_DYNAMIC) continue;
    const segment = { address: u64(header + 16), offset: u64(header + 8), size: u64(header + 32) };
    requireRange(Number(segment.offset), Number(segment.size), 'segment');
    if (type === PT_LOAD) loads.push(segment);
    else dynamic = segment;
  }
  if (!dynamic) return;

  const mapAddress = (address, label) => {
    for (const load of loads) {
      const delta = address - load.address;
      if (delta >= 0n && delta < load.size) {
        return { offset: Number(load.offset + delta), end: Number(load.offset + load.size) };
      }
    }
    throw new Error(`ELF ${label} address is outside file-backed PT_LOAD segments`);
  };

  const tags = new Map();
  const dynamicEnd = Number(dynamic.offset + dynamic.size);
  for (let entry = Number(dynamic.offset); ; entry += 16) {
    requireRange(entry, 16, 'dynamic table', dynamicEnd);
    const tag = u64(entry);
    if (tag === DT_NULL) break;
    tags.set(tag, u64(entry + 8));
  }

  const stringReader = (address) => {
    const strings = mapAddress(address, 'DT_STRTAB');
    const size = Number(tags.get(DT_STRSZ) ?? (strings.end - strings.offset));
    requireRange(strings.offset, size, 'dynamic string table', strings.end);
    const table = bytes.subarray(strings.offset, strings.offset + size);
    const decoder = new TextDecoder();
    return (start, label) => {
      const end = table.indexOf(0, start);
      if (end === -1) throw new Error(`invalid ELF ${label}`);
      return decoder.decode(table.subarray(start, end));
    };
  };

  return { tags, u16, u32, requireRange, mapAddress, stringReader };
}

/** @param {Uint8Array} bytes @returns {string[]} */
export function requiredGlibcVersions(bytes) {
  const dynamic = readDynamic(bytes);
  if (!dynamic) return [];
  const { tags, u16, u32, requireRange, mapAddress, stringReader } = dynamic;
  const needAddress = tags.get(DT_VERNEED);
  const countTag = tags.get(DT_VERNEEDNUM);
  if (needAddress === undefined && countTag === undefined) return [];
  const stringAddress = tags.get(DT_STRTAB);
  if (needAddress === undefined || countTag === undefined || stringAddress === undefined) {
    throw new Error('ELF version needs require DT_VERNEED, DT_VERNEEDNUM and DT_STRTAB');
  }
  if (countTag === 0n) return [];

  const readString = stringReader(stringAddress);
  const needs = mapAddress(needAddress, 'DT_VERNEED');
  const needCount = Number(countTag);
  const versions = new Set();
  for (let index = 0, need = needs.offset; index < needCount; index += 1) {
    requireRange(need, VERNEED_SIZE, 'Verneed', needs.end);
    const count = u16(need + 2);
    const auxOffset = u32(need + 8);
    if (count === 0) throw new Error('invalid ELF vn_cnt');
    if (auxOffset < VERNEED_SIZE) throw new Error('invalid ELF vn_aux offset');
    for (let item = 0, aux = need + auxOffset; item < count; item += 1) {
      requireRange(aux, VERNAUX_SIZE, 'Vernaux', needs.end);
      const name = readString(u32(aux + 8), 'Vernaux name');
      if (/^GLIBC_\d+\.\d+(?:\.\d+)?$/.test(name)) versions.add(name);
      else if (name.startsWith('GLIBC_')) throw new Error(`unsupported glibc version need ${name}`);

      const next = u32(aux + 12);
      if (item + 1 < count ? next < VERNAUX_SIZE : next !== 0) {
        throw new Error('invalid ELF vna_next chain');
      }
      aux += next;
    }
    const next = u32(need + 12);
    if (index + 1 < needCount ? next < VERNEED_SIZE : next !== 0) {
      throw new Error('invalid ELF vn_next chain');
    }
    need += next;
  }
  return [...versions].sort(compareGlibc);
}

function dynamicSymbolCount({ tags, u32, requireRange, mapAddress }) {
  const hashAddress = tags.get(DT_HASH);
  if (hashAddress !== undefined) {
    const hash = mapAddress(hashAddress, 'DT_HASH');
    requireRange(hash.offset, 8, 'DT_HASH header', hash.end);
    return u32(hash.offset + 4);
  }

  const gnuHashAddress = tags.get(DT_GNU_HASH);
  if (gnuHashAddress === undefined) throw new Error('ELF dynamic symbols require DT_HASH or DT_GNU_HASH');
  const hash = mapAddress(gnuHashAddress, 'DT_GNU_HASH');
  requireRange(hash.offset, 16, 'DT_GNU_HASH header', hash.end);
  const bucketCount = u32(hash.offset);
  const symbolOffset = u32(hash.offset + 4);
  const bloomSize = u32(hash.offset + 8);
  const buckets = hash.offset + 16 + bloomSize * 8;
  requireRange(buckets, bucketCount * 4, 'DT_GNU_HASH buckets', hash.end);
  let lastBucketStart = 0;
  for (let index = 0; index < bucketCount; index += 1) {
    lastBucketStart = Math.max(lastBucketStart, u32(buckets + index * 4));
  }
  if (lastBucketStart === 0) throw new Error('ELF DT_GNU_HASH lists no symbols, so the dynamic symbol count is unknown');
  if (lastBucketStart < symbolOffset) throw new Error('invalid ELF DT_GNU_HASH bucket');

  const chains = buckets + bucketCount * 4;
  for (let index = lastBucketStart; ; index += 1) {
    const chain = chains + (index - symbolOffset) * 4;
    requireRange(chain, 4, 'DT_GNU_HASH chain', hash.end);
    if (u32(chain) & 1) return index + 1;
  }
}

/** @param {Uint8Array} bytes @returns {string[]} */
export function unversionedUndefinedSymbols(bytes) {
  const dynamic = readDynamic(bytes);
  if (!dynamic) return [];
  const { tags, u16, u32, requireRange, mapAddress, stringReader } = dynamic;
  const symbolAddress = tags.get(DT_SYMTAB);
  if (symbolAddress === undefined) return [];
  if (Number(tags.get(DT_SYMENT)) !== SYM_SIZE) throw new Error('invalid ELF DT_SYMENT');
  const stringAddress = tags.get(DT_STRTAB);
  if (stringAddress === undefined) throw new Error('ELF dynamic symbols require DT_STRTAB');

  const count = dynamicSymbolCount(dynamic);
  const symbols = mapAddress(symbolAddress, 'DT_SYMTAB');
  requireRange(symbols.offset, count * SYM_SIZE, 'dynamic symbol table', symbols.end);
  const versionAddress = tags.get(DT_VERSYM);
  const versions = versionAddress === undefined ? undefined : mapAddress(versionAddress, 'DT_VERSYM');
  if (versions) requireRange(versions.offset, count * 2, 'symbol version table', versions.end);

  const readString = stringReader(stringAddress);
  const names = [];
  for (let index = 0; index < count; index += 1) {
    const symbol = symbols.offset + index * SYM_SIZE;
    const binding = bytes[symbol + 4] >>> 4;
    const section = u16(symbol + 6);
    if (section !== SHN_UNDEF || binding === STB_WEAK) continue;
    if (versions && (u16(versions.offset + index * 2) & 0x7fff) > 1) continue;
    const name = readString(u32(symbol), 'dynamic symbol name');
    if (name && !/^(napi_|node_api_|uv_)/.test(name)) names.push(name);
  }
  return names;
}

function main(files) {
  if (files.length === 0) {
    console.error('usage: node scripts/glibc-floor.mjs <file.node>...');
    process.exitCode = 1;
  }
  for (const file of files) {
    try {
      const bytes = readFileSync(file);
      for (const name of unversionedUndefinedSymbols(bytes)) {
        console.error(`${file}: unversioned undefined symbol ${name}`);
        process.exitCode = 1;
      }
      const version = requiredGlibcVersions(bytes).at(-1);
      if (!version) throw new Error('no GLIBC version needs found');
      const max = version.slice('GLIBC_'.length);
      console.log(`${file}: needs GLIBC ${max} (floor ${GLIBC_FLOOR})`);
      if (compareGlibc(max, GLIBC_FLOOR) > 0) process.exitCode = 1;
    } catch (error) {
      console.error(`${file}: ${error.message}`);
      process.exitCode = 1;
    }
  }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) main(process.argv.slice(2));
