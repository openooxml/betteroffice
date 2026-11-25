import { expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { GLIBC_FLOOR, compareGlibc, requiredGlibcVersions } from './glibc-floor.mjs';

function elf(libraries: string[][], machine = 62) {
  let names = '\0';
  const addName = (name: string) => {
    const offset = names.length;
    names += `${name}\0`;
    return offset;
  };
  const needs = libraries.map((versions, index) => ({
    file: addName(index === 0 ? 'libc.so.6' : 'libm.so.6'),
    versions: versions.map(addName)
  }));

  const loadAddress = 0x400000;
  const elfHeaderSize = 64;
  const headersOffset = elfHeaderSize;
  const headerSize = 56;
  const dynamicOffset = headersOffset + 2 * headerSize;
  const stringsOffset = dynamicOffset + 5 * 16;
  const needsOffset = stringsOffset + Math.ceil(names.length / 8) * 8;
  const recordStride = 24;
  const size = needsOffset + needs.reduce((sum, { versions }) => sum + 1 + versions.length, 0) * recordStride;
  const bytes = new Uint8Array(size + 8).subarray(8);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = (offset: number, value: number) => view.setUint16(offset, value, true);
  const u32 = (offset: number, value: number) => view.setUint32(offset, value, true);
  const u64 = (offset: number, value: number) => view.setBigUint64(offset, BigInt(value), true);

  bytes.set([0x7f, 0x45, 0x4c, 0x46, 2, 1, 1]);
  u16(16, 3);
  u16(18, machine);
  u32(20, 1);
  u64(32, headersOffset);
  u16(52, elfHeaderSize);
  u16(54, headerSize);
  u16(56, 2);

  const dynamic: [number, number][] = [
    [5, loadAddress + stringsOffset],
    [10, names.length]
  ];
  if (needs.length) dynamic.push([0x6ffffffe, loadAddress + needsOffset], [0x6fffffff, needs.length]);
  dynamic.push([0, 0]);
  function programHeader(header: number, type: number, offset: number, segmentSize: number, align: number) {
    u32(header, type);
    u32(header + 4, 4);
    u64(header + 8, offset);
    u64(header + 16, loadAddress + offset);
    u64(header + 32, segmentSize);
    u64(header + 40, segmentSize);
    u64(header + 48, align);
  }
  programHeader(headersOffset, 1, 0, bytes.length, 0x1000);
  programHeader(headersOffset + headerSize, 2, dynamicOffset, dynamic.length * 16, 8);
  dynamic.forEach(([tag, value], index) => {
    const offset = dynamicOffset + index * 16;
    u64(offset, tag);
    u64(offset + 8, value);
  });
  bytes.set(new TextEncoder().encode(names), stringsOffset);

  let need = needsOffset;
  needs.forEach(({ file, versions }, index) => {
    const next = (1 + versions.length) * recordStride;
    u16(need, 1);
    u16(need + 2, versions.length);
    u32(need + 4, file);
    u32(need + 8, recordStride);
    u32(need + 12, index + 1 < needs.length ? next : 0);
    versions.forEach((name, item) => {
      const aux = need + (item + 1) * recordStride;
      u32(aux + 8, name);
      u32(aux + 12, item + 1 < versions.length ? recordStride : 0);
    });
    need += next;
  });
  return bytes;
}

const cli = (...files: string[]) => spawnSync('node', [
  fileURLToPath(new URL('./glibc-floor.mjs', import.meta.url)),
  ...files
], { encoding: 'utf8' });

test('collects sorted unique glibc versions from every needed library', () => {
  for (const machine of [62, 183]) {
    expect(requiredGlibcVersions(elf([
      ['GLIBC_2.43', 'GLIBC_2.9', 'GLIBC_PRIVATE'],
      [`GLIBC_${GLIBC_FLOOR}`, 'GCC_3.0', 'CXXABI_1.3', 'GLIBC_2.9', 'GLIBC_2.2.5']
    ], machine))).toEqual(['GLIBC_2.2.5', 'GLIBC_2.9', `GLIBC_${GLIBC_FLOOR}`, 'GLIBC_2.43']);
  }
});

test('compares major, minor and patch versions numerically', () => {
  expect(compareGlibc('2.9', GLIBC_FLOOR)).toBeLessThan(0);
  expect(compareGlibc(GLIBC_FLOOR, '2.43')).toBeLessThan(0);
  expect(compareGlibc('2.43.10', '2.43.9')).toBeGreaterThan(0);
  expect(compareGlibc(`${GLIBC_FLOOR}.0`, GLIBC_FLOOR)).toBe(0);
});

test('rejects non-ELF, ELF32, big-endian and other ELF machines', () => {
  expect(() => requiredGlibcVersions(new Uint8Array(64))).toThrow('expected an ELF file');
  const elf32 = elf([]);
  elf32[4] = 1;
  expect(() => requiredGlibcVersions(elf32)).toThrow('expected little-endian ELF64');
  const bigEndian = elf([]);
  bigEndian[5] = 2;
  expect(() => requiredGlibcVersions(bigEndian)).toThrow('expected little-endian ELF64');
  expect(() => requiredGlibcVersions(elf([], 243))).toThrow('expected x86_64 or aarch64 ELF');
});

test('the CLI checks every file, rejecting newer binaries and files without version needs', () => {
  const directory = mkdtempSync(join(tmpdir(), 'glibc-floor-'));
  try {
    const compatible = join(directory, 'compatible.node');
    const newer = join(directory, 'newer.node');
    const noNeeds = join(directory, 'no-needs.node');
    writeFileSync(compatible, elf([[`GLIBC_${GLIBC_FLOOR}`], ['GLIBC_2.9']]));
    writeFileSync(newer, elf([[`GLIBC_${GLIBC_FLOOR}`], ['GLIBC_2.43']]));
    writeFileSync(noNeeds, elf([]));

    const pass = cli(compatible);
    expect(pass.status).toBe(0);
    const compatibleLine = `${compatible}: needs GLIBC ${GLIBC_FLOOR} (floor ${GLIBC_FLOOR})\n`;
    expect(pass.stdout).toBe(compatibleLine);

    const fail = cli(newer, compatible);
    expect(fail.status).toBe(1);
    expect(fail.stdout).toBe(`${newer}: needs GLIBC 2.43 (floor ${GLIBC_FLOOR})\n${compatibleLine}`);

    const missing = cli(noNeeds);
    expect(missing.status).toBe(1);
    expect(missing.stderr).toContain(`${noNeeds}: no GLIBC version needs found`);
    expect(cli(join(directory, 'missing.*.node')).status).toBe(1);
    expect(cli().status).toBe(1);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
