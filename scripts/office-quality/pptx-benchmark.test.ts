import { expect, test } from 'bun:test';
import { mergePptxBenchmark, METHOD } from './pptx-benchmark.mjs';
import { renderSection } from './readme.mjs';

const hash = 'a'.repeat(64);
const commit = 'b'.repeat(40);
const samples = [{id:'deck',format:'pptx',metadata:{source:{sha256:hash},reference:{pages:2}},comparisons:[]}];
const report = () => ({source_sha:commit,commit,versions:{docx:'1.0.0',pptx:'1.0.0',xlsx:'1.0.0'},samples});
const plan = () => ({samples,source_sha:commit,pptx_published_source_sha:'c'.repeat(40),versions:{pptx:'1.0.0'}});
const build = (source_sha:string) => ({source_sha,binary_sha256:hash,harness_sha256:hash,rustc:'rustc test',profile:'release'});
const part = () => ({schema_version:1,plan_sha256:hash,
  benchmark:{method:METHOD,dpi:96,trials:5,published_version:'1.0.0',libreoffice_version:'26.2.3.2',fonts_sha256:hash,
    builds:{published:build('c'.repeat(40)),commit:build(commit)}},
  samples:[{id:'deck',source_sha256:hash,timings:Object.fromEntries(['published','commit','libreoffice'].map(channel => [channel,{status:'ok',elapsed_ms:[100,100,100,100,100]}])),libreoffice:{channel:'libreoffice',version:'26.2.3.2',status:'ok',
    source_verified:true,reference:{status:'ok',sha256:hash},actual:{status:'ok',sha256:hash},
    penalized_ssim:0.75,reference_pages:2,actual_pages:2,resized:false}}],
});

test('adds a versioned PPTX LibreOffice comparison with scored coverage', () => {
  const merged = mergePptxBenchmark(plan(),report(),part(),hash);
  const text = renderSection(merged).split('### PPTX')[1].split('### XLSX')[0];
  expect(text).toContain('LibreOffice (26.2.3.2)');
  expect(text).toContain('<td align="right">0.7500</td>');
  expect(text).toContain('<td align="right">1/1</td>');
  expect(text).toContain('Render time (avg)');
  expect(text).toContain('100 ms');
  expect(text).not.toContain('Exact page counts');
  expect(text).not.toContain('Absolute page error');
});

test('keeps PPTX export failures and refuses incomplete or mismatched reports', () => {
  const failed = part();
  failed.samples[0].libreoffice = {channel:'libreoffice',version:'26.2.3.2',status:'failed',stage:'capture',error:'export failed'} as any;
  const merged = mergePptxBenchmark(plan(),report(),failed,hash);
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
    expect(() => mergePptxBenchmark(plan(),report(),input,hash)).toThrow();
  }
});
