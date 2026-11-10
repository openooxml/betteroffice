import { expect, test } from 'bun:test';
import { mergePptxBenchmark, pptxShards, METHOD } from './pptx-benchmark.mjs';
import { renderSection } from './readme.mjs';

const hash = 'a'.repeat(64);
const commit = 'b'.repeat(40);
const samples = [{id:'deck',format:'pptx',metadata:{source:{sha256:hash},reference:{pages:2}},comparisons:[]}];
const report = () => ({source_sha:commit,commit,versions:{docx:'1.0.0',pptx:'1.0.0',xlsx:'1.0.0'},samples});
const plan = () => ({samples,source_sha:commit,pptx_published_source_sha:'c'.repeat(40),versions:{pptx:'1.0.0'}});
const build = (source_sha:string) => ({source_sha,binary_sha256:hash,harness_sha256:hash,rustc:'rustc test',profile:'release'});
const part = () => ({schema_version:1,plan_sha256:hash,shard:0,
  benchmark:{method:METHOD,dpi:96,trials:5,published_version:'1.0.0',libreoffice_version:'26.2.3.2',fonts_sha256:hash,
    builds:{published:build('c'.repeat(40)),commit:build(commit)}},
  samples:[{id:'deck',source_sha256:hash,timings:Object.fromEntries(['published','commit','libreoffice'].map(channel => [channel,{status:'ok',elapsed_ms:[100,100,100,100,100]}])),libreoffice:{channel:'libreoffice',version:'26.2.3.2',status:'ok',
    source_verified:true,reference:{status:'ok',sha256:hash},actual:{status:'ok',sha256:hash},
    penalized_ssim:0.75,reference_pages:2,actual_pages:2,resized:false}}],
});

test('adds a versioned PPTX LibreOffice comparison with scored coverage', () => {
  const merged = mergePptxBenchmark(plan(),report(),[part()],hash);
  const text = renderSection(merged).split('### PPTX')[1].split('### XLSX')[0];
  expect(text).toContain('LibreOffice (26.2.3.2)');
  expect(text).toContain('<td align="right">0.7500</td>');
  expect(text).toContain('<td align="right">1/1</td>');
  expect(text).toContain('Render time (avg)');
  expect(text).toContain('100 ms');
  expect(text).toContain('<tr><td>Exact page counts</td><td align="right">—</td><td align="right">—</td><td align="right">1/1</td></tr>');
  expect(text).toContain('<tr><td>Absolute page error</td><td align="right">—</td><td align="right">—</td><td align="right">0</td></tr>');
});

test('keeps PPTX export failures and refuses incomplete or mismatched reports', () => {
  const failed = part();
  failed.samples[0].libreoffice = {channel:'libreoffice',version:'26.2.3.2',status:'failed',stage:'capture',error:'export failed'} as any;
  const merged = mergePptxBenchmark(plan(),report(),[failed],hash);
  expect(merged.samples[0].comparisons[0].status).toBe('failed');
  for (const mutate of [
    (p:any) => p.samples[0].timings.commit.elapsed_ms.pop(),
    (p:any) => p.benchmark.builds.commit.source_sha='d'.repeat(40),
    (p:any) => p.benchmark.builds.published.harness_sha256='d'.repeat(64),
    (p:any) => p.samples.pop(),
    (p:any) => p.samples.push(p.samples[0]),
    (p:any) => p.plan_sha256='c'.repeat(64),
    (p:any) => p.samples[0].libreoffice.reference_pages=3,
    (p:any) => p.samples[0].libreoffice.actual.sha256='c'.repeat(64),
  ]) {
    const input = part();
    mutate(input);
    expect(() => mergePptxBenchmark(plan(),report(),[input],hash)).toThrow();
  }
});

test('a presentation set too large for one job is split into budgeted shards, all of which must be reported', () => {
  const decks = Array.from({length:60},(_,index) => ({id:`deck-${String(index).padStart(2,'0')}`,format:'pptx',
    metadata:{source:{sha256:hash},reference:{pages:100}},comparisons:[]}));
  const big = () => ({...plan(),samples:decks});
  const bigReport = () => ({...report(),samples:decks});
  const shards = pptxShards(big());
  expect(shards.length).toBeGreaterThan(1);
  const parts = () => shards.map((ids:string[],shard:number) => ({...part(),shard,
    samples:ids.map(id => ({...part().samples[0],id,libreoffice:{...part().samples[0].libreoffice,reference_pages:100,actual_pages:100}}))}));
  const merged = mergePptxBenchmark(big(),bigReport(),parts(),hash);
  expect(merged.samples.every((sample:any) => sample.comparisons.length === 1 && sample.native_timings)).toBe(true);
  expect(() => mergePptxBenchmark(big(),bigReport(),parts().slice(1),hash)).toThrow('Missing or unexpected PPTX benchmark shards');
  const duplicate = parts(); duplicate[1].shard = 0;
  expect(() => mergePptxBenchmark(big(),bigReport(),duplicate,hash)).toThrow('Duplicate or unknown PPTX benchmark shard');
  const differing = parts(); differing[1].benchmark = {...differing[1].benchmark,fonts_sha256:'f'.repeat(64)};
  expect(() => mergePptxBenchmark(big(),bigReport(),differing,hash)).toThrow('PPTX benchmark shard environments differ');
  const swapped = parts(); [swapped[0].samples[0], swapped[1].samples[0]] = [swapped[1].samples[0], swapped[0].samples[0]];
  expect(() => mergePptxBenchmark(big(),bigReport(),swapped,hash)).toThrow('Invalid PPTX LibreOffice sample identity');
});
