import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { expect, test } from 'bun:test';
import JSZip from 'jszip';

const MiB = 2 ** 20;
const budget = 256 * MiB;
const minimumGrowth = budget / 2;
const skipReason = process.platform !== 'linux'
  ? 'requires Linux RSS accounting and glibc malloc_trim'
  : !/\/(?:libc\.so\.6|libc-\d[\d.]*\.so)(?:\s|$)/m.test(readFileSync('/proc/self/maps', 'utf8'))
    ? 'requires glibc for the System allocator release path'
    : undefined;

async function numbersWorkbook(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  zip.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
    + '<row><c><v>1</v></c><c><v>2</v></c></row>'.repeat(1_000_000)
    + '</sheetData></worksheet>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 1 } });
}

type Measurement = {
  failure: string | null;
  anchors: number;
  rssBefore: number;
  rssAfter: number;
  peakBefore: number;
  peakAfter: number;
};

test.skipIf(skipReason !== undefined)(`post-failure release returns at least half the RSS growth${skipReason ? ` (skipped: ${skipReason})` : ''}`, async () => {
  const workbook = await numbersWorkbook();
  const bindingPath = fileURLToPath(new URL('./index.js', import.meta.url));
  const child = spawnSync(process.execPath, ['--eval', `
    const { readFileSync } = require('node:fs');
    const native = require(${JSON.stringify(bindingPath)});
    const workbook = new Uint8Array(readFileSync(0));
    const anchors = [];
    let nextAnchor = native.engineHeapLiveBytes() + ${16 * MiB};
    let document;
    let failure = null;
    native.setHeapLimit(${budget});
    try {
      globalThis.__betterofficeEngineProgress = () => {
        const live = native.engineHeapLiveBytes();
        if (live >= nextAnchor) {
          nextAnchor = live + ${16 * MiB};
          anchors.push(new native.XlsxSnapshotBuilder());
        }
      };
      const rssBefore = process.memoryUsage().rss;
      const peakBefore = process.resourceUsage().maxRSS * 1024;
      try {
        document = native.XlsxDocument.open(workbook);
      } catch (error) {
        failure = error instanceof Error ? error.message : String(error);
      }
      const rssAfter = process.memoryUsage().rss;
      const peakAfter = process.resourceUsage().maxRSS * 1024;
      process.stdout.write(JSON.stringify({ failure, anchors: anchors.length, rssBefore, rssAfter, peakBefore, peakAfter }));
    } finally {
      native.setHeapLimit(0);
      delete globalThis.__betterofficeEngineProgress;
      document?.free();
      for (const anchor of anchors) anchor.free();
    }
  `], { input: workbook, encoding: 'utf8', timeout: 90_000 });

  const diagnostics = JSON.stringify({ status: child.status, signal: child.signal, error: child.error?.message, stdout: child.stdout, stderr: child.stderr });
  expect(child.error, diagnostics).toBeUndefined();
  expect(child.status, diagnostics).toBe(0);
  const measurement = JSON.parse(child.stdout) as Measurement;
  const { failure, anchors, rssBefore, rssAfter, peakBefore, peakAfter } = measurement;
  const growth = peakAfter - rssBefore;
  const returned = peakAfter - rssAfter;
  const measured = JSON.stringify({ ...measurement, growth, returned });

  expect(failure, measured).toContain('BetterOffice engine memory budget exceeded');
  expect(anchors, measured).toBeGreaterThanOrEqual(4);
  expect(peakAfter - peakBefore, measured).toBeGreaterThanOrEqual(minimumGrowth);
  expect(growth, measured).toBeGreaterThanOrEqual(minimumGrowth);
  expect(returned, measured).toBeGreaterThanOrEqual(growth / 2);
}, 120_000);
