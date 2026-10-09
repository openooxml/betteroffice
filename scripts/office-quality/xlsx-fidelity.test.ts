import { expect, test } from 'bun:test';
import { mergeXlsxFidelity, METHOD, xlsxFidelityShards } from './xlsx-fidelity.mjs';
import { renderSection } from './readme.mjs';

const hash = 'a'.repeat(64);
const commit = 'b'.repeat(40);
const samples = [0, 1, 2, 3, 4].map(index => ({id:`book-${index}`,format:'xlsx',
  metadata:{source:{sha256:hash},reference:{pages:index + 1}},comparisons:[]}));
const plan = () => ({samples,source_sha:commit});
const report = () => ({...plan(),commit,versions:{docx:'1.0.0',pptx:'1.0.0',xlsx:'1.0.0'}});
const parts = () => xlsxFidelityShards(plan()).map((ids, shard) => ({schema_version:1,plan_sha256:hash,shard,
  benchmark:{method:METHOD,dpi:150,source_sha:commit,libreoffice_version:'26.2.3.2',fonts_sha256:hash},
  samples:ids.map(id => ({id,source_sha256:hash,libreoffice:{channel:'libreoffice',version:'26.2.3.2',status:'ok',
    source_verified:true,reference:{status:'ok',sha256:hash},actual:{status:'ok',sha256:hash},
    penalized_ssim:0.75,reference_pages:samples.find(s => s.id === id)!.metadata.reference.pages,
    actual_pages:samples.find(s => s.id === id)!.metadata.reference.pages,resized:false}})),
}));

test('XLSX fidelity includes non-formula workbooks and balances recorded ranges', () => {
  const shards = xlsxFidelityShards(plan());
  expect(shards).toHaveLength(4);
  expect(shards.flat().sort()).toEqual(samples.map(s => s.id));
  expect(xlsxFidelityShards({samples:[]})).toEqual([]);
});

test('renders real LibreOffice XLSX SSIM and scored coverage without calculation references', () => {
  const inputs = parts();
  inputs[0].samples[0].libreoffice = {channel:'libreoffice',version:'26.2.3.2',status:'failed',stage:'capture',error:'overflow'} as any;
  const merged = mergeXlsxFidelity(plan(), report(), inputs, hash);
  const text = renderSection(merged).split('### XLSX')[1];
  expect(text).toContain('LibreOffice (26.2.3.2)');
  expect(text).toContain('<td align="right">0.7500</td>');
  expect(text).toContain('<td align="right">4/5</td>');
  expect(text).not.toContain('Recalc accuracy');
  expect(text).not.toContain('Exact page counts');
});

test('rejects incomplete, mixed, duplicate, or incorrectly assigned shards', () => {
  for (const mutate of [
    (p:any) => p.pop(), (p:any) => p[0].samples.pop(),
    (p:any) => p[0].shard = p[1].shard,
    (p:any) => p[0].plan_sha256 = 'c'.repeat(64),
    (p:any) => p[0].benchmark.source_sha = 'c'.repeat(40),
    (p:any) => p[0].benchmark.fonts_sha256 = 'c'.repeat(64),
    (p:any) => p[0].samples[0].id = p[1].samples[0].id,
    (p:any) => p[0].samples[0].libreoffice.actual_pages++,
    (p:any) => p[0].samples[0].libreoffice.actual.sha256 = 'c'.repeat(64),
  ]) {
    const inputs = parts();
    mutate(inputs);
    expect(() => mergeXlsxFidelity(plan(), report(), inputs, hash)).toThrow();
  }
  expect(() => mergeXlsxFidelity(plan(), {...report(),xlsx_benchmark:{libreoffice_version:'25.2.0.0'}}, parts(), hash)).toThrow();
});
