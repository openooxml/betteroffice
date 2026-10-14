import { expect, test } from 'bun:test';
import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { createPlan } from './plan.mjs';
import { writeCachedAsset } from './asset-cache.mjs';
import { mergeRoundtrips, METHOD, roundtripSummary, roundtripShards, SHARD_SIZE, PARALLELISM, ENGINE_TIMEOUT, HELPER_TIMEOUT } from './roundtrip.mjs';
import { renderSection } from './readme.mjs';

const hash = 'a'.repeat(64);
const sha = 'b'.repeat(40);
const published = 'c'.repeat(40);
const formats = ['docx', 'pptx', 'xlsx'];
const samples = formats.flatMap(format => ['one', 'two'].map(name => ({ id:`${format}-${name}`, format,
  metadata:{source:{sha256:hash}}, comparisons:[] })));
const plan = () => ({ samples, formats, source_sha:sha, commit:sha,
  versions:Object.fromEntries(formats.map(format => [format, '1.0.0'])),
  ...Object.fromEntries(formats.map(format => [`${format}_published_source_sha`, published])) });
const build = (source_sha:string) => ({source_sha,binary_sha256:hash,harness_sha256:hash,rustc:'rustc test',profile:'release'});
const success = () => ({parse:{status:'ok'},native:{parse:'ok',source_sha256:hash,stage:'complete',edit_verified:true,output_sha256:hash},
  roundtrip:{status:'ok',stage:'preserve',edit_matches:true,original_parts:3,identical_parts:2,
    added_parts:[],removed_parts:[],changed_parts:['document.xml'],unrelated_changed_parts:[]}});
const parts = () => formats.map(format => ({schema_version:1,plan_sha256:hash,format,shard:0,
  benchmark:{method:METHOD,published_version:'1.0.0',checker_sha256:hash,timeout_seconds:ENGINE_TIMEOUT,helper_timeout_seconds:HELPER_TIMEOUT,
    parallelism:PARALLELISM,shard_size:SHARD_SIZE,libreoffice_version:'26.2.3.2',libreoffice_build:'LibreOffice 26.2.3.2',libreoffice_host_sha256:hash,
    builds:{published:build(published),commit:build(sha)}},
  samples:samples.filter(s => s.format === format).map(s => ({id:s.id,source_sha256:hash,
    probe:{status:'ok',part:'document.xml',old:'before',new:'after'},
    channels:{published:success(),commit:success(),libreoffice:success()}})),
}));

test('parse and preservation use all planned files and remain separate after save failures', () => {
  const input = parts();
  input[0].samples[0].channels.commit.roundtrip = {status:'failed',stage:'save',error:'write failed'} as any;
  input[0].samples[1].channels.published = {parse:{status:'failed',error:'cannot parse'},roundtrip:{status:'failed',stage:'parse',error:'cannot parse'}} as any;
  const merged = mergeRoundtrips(plan(), plan(), input, hash);
  expect(roundtripSummary(merged.samples, 'docx')).toEqual({
    published:{parsed:1,preserved:1,total:2}, commit:{parsed:2,preserved:1,total:2}, libreoffice:{parsed:2,preserved:2,total:2},
  });
  const text = renderSection(merged).split('### DOCX')[1].split('### PPTX')[0];
  expect(text).toContain('<td>Parse success</td><td align="right">50.00%</td><td align="right">100.00%</td>');
  expect(text).not.toContain('Lossless roundtrip');
});

test('requires exact sample/build identities and proof of a nonempty preserved edit', () => {
  for (const mutate of [
    (p:any) => p.pop(), (p:any) => p[0].samples.pop(), (p:any) => p[0].format = 'xlsx',
    (p:any) => p[0].samples[1].id = p[0].samples[0].id,
    (p:any) => p[0].plan_sha256 = 'c'.repeat(64),
    (p:any) => p[0].benchmark.builds.commit.source_sha = published,
    (p:any) => p[0].benchmark.checker_sha256 = 'c'.repeat(64),
    (p:any) => p[0].benchmark.builds.commit.harness_sha256 = 'c'.repeat(64),
    (p:any) => p[0].samples[0].channels.commit.native.source_sha256 = 'c'.repeat(64),
    (p:any) => p[0].samples[0].channels.commit.native.edit_verified = false,
    (p:any) => p[0].samples[0].channels.commit.roundtrip.unrelated_changed_parts.push('custom.xml'),
    (p:any) => p[0].samples[0].channels.commit.roundtrip.removed_parts.push('image.png'),
    (p:any) => p[0].samples[0].probe.new = 'before',
    (p:any) => p[0].samples[0].probe.status = 'unavailable',
    (p:any) => delete p[0].samples[0].channels.libreoffice,
    (p:any) => p[0].shard = 1,
    (p:any) => p[0].benchmark.parallelism = 1,
  ]) {
    const input = parts();
    mutate(input);
    expect(() => mergeRoundtrips(plan(), plan(), input, hash)).toThrow();
  }
});


test('LibreOffice parse/preservation is measured and stale results are rejected', () => {
  const merged = mergeRoundtrips(plan(), plan(), parts(), hash);
  const report = {...merged,xlsx_fidelity_benchmark:{source_sha:sha,libreoffice_version:'26.2.3.2'}};
  const text = renderSection(report).split('### XLSX')[1];
  expect(text).toContain('<td>Parse success</td><td align="right">100.00%</td><td align="right">100.00%</td><td align="right">100.00%</td>');
  expect(renderSection(report)).not.toContain('Lossless roundtrip');
  expect(() => renderSection({...merged,source_sha:published})).toThrow('revision');
});


test('even 1000 files fit bounded shards and leave time for setup and artifact upload', () => {
  const many = {...plan(), samples:Array.from({length:1000}, (_, i) => ({...samples[0],id:`file-${i}`}))};
  const shards = roundtripShards(many);
  expect(shards.length).toBe(63);
  expect(shards.flatMap(shard => shard.ids).sort()).toEqual(many.samples.map(sample => sample.id).sort());
  expect(shards.every(shard => shard.ids.length <= SHARD_SIZE)).toBe(true);
  const worstSeconds = Math.ceil(SHARD_SIZE / PARALLELISM) * (3 * (ENGINE_TIMEOUT + HELPER_TIMEOUT) + HELPER_TIMEOUT);
  expect(worstSeconds).toBe(44 * 60);
  expect(worstSeconds).toBeLessThan(60 * 60);
});

test('all shards of a format must be present and share the same environment', () => {
  const input = {...plan(),samples:[...samples]};
  input.samples.push(...Array.from({length:16}, (_, i) => ({...samples[0],id:`docx-extra-${i}`})));
  const byFormat = Object.fromEntries(parts().map(part => [part.format, part]));
  const reports = roundtripShards(input).map(({format,shard,ids}) => ({...byFormat[format],shard,
    samples:ids.map(id => ({...structuredClone(byFormat[format].samples[0]),id}))}));
  const merged = mergeRoundtrips(input, input, reports, hash);
  expect(roundtripSummary(merged.samples,'docx').libreoffice.total).toBe(18);
  expect(() => mergeRoundtrips(input, input, reports.slice(1), hash)).toThrow();
  reports[1].benchmark = {...reports[1].benchmark,libreoffice_build:'different build'};
  expect(() => mergeRoundtrips(input, input, reports, hash)).toThrow();
});


test('the CLI accepts a string shard from Actions and stages only that shard from verified cached fixtures', async () => {
  const root = await mkdtemp(join(tmpdir(), 'roundtrip-stage-test-'));
  try {
    const bytes = Buffer.from('synthetic staging fixture');
    const sha256 = createHash('sha256').update(bytes).digest('hex');
    const source = {url:'https://corpus.betteroffice.dev/docx-one/source.docx',bytes:bytes.length,sha256};
    const fixture = createPlan({...plan(),react_version:'1.0.0',samples:[{id:'docx-one',format:'docx',metadata:{
      format:'docx',source,reference:{status:'ok',dpi:150,sha256,pages:1},
      reference_pages:[{url:'https://corpus.betteroffice.dev/docx-one/reference/page_0001.png',bytes:1,sha256}],
    }}]});
    const planPath = join(root,'plan.json');
    await writeFile(planPath,JSON.stringify(fixture));
    await writeCachedAsset(join(root,'cache'),source,bytes);
    await promisify(execFile)('node',[fileURLToPath(new URL('./roundtrip.mjs',import.meta.url))],{env:{...process.env,
      QUALITY_PLAN:planPath,QUALITY_FORMAT:'docx',QUALITY_SHARD:'0',QUALITY_OUTPUT:join(root,'output'),QUALITY_ASSET_CACHE:join(root,'cache'),
    }});
    const job = JSON.parse(await readFile(join(root,'output/job.json'),'utf8'));
    expect(job.shard).toBe(0);
    expect(job.samples.map((sample:any) => sample.id)).toEqual(['docx-one']);
    expect(await readFile(join(root,'output/docx-one/source.docx'))).toEqual(bytes);
  } finally {
    await rm(root,{recursive:true,force:true});
  }
});
