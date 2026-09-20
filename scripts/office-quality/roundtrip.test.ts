import { expect, test } from 'bun:test';
import { mergeRoundtrips, METHOD, roundtripSummary } from './roundtrip.mjs';
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
const parts = () => formats.map(format => ({schema_version:1,plan_sha256:hash,format,
  benchmark:{method:METHOD,published_version:'1.0.0',checker_sha256:hash,timeout_seconds:180,
    builds:{published:build(published),commit:build(sha)}},
  samples:samples.filter(s => s.format === format).map(s => ({id:s.id,source_sha256:hash,
    probe:{status:'ok',part:'document.xml',old:'before',new:'after'},
    channels:{published:success(),commit:success()}})),
}));

test('parse and preservation use all planned files and remain separate after save failures', () => {
  const input = parts();
  input[0].samples[0].channels.commit.roundtrip = {status:'failed',stage:'save',error:'write failed'} as any;
  input[0].samples[1].channels.published = {parse:{status:'failed',error:'cannot parse'},roundtrip:{status:'failed',stage:'parse',error:'cannot parse'}} as any;
  const merged = mergeRoundtrips(plan(), plan(), input, hash);
  expect(roundtripSummary(merged.samples, 'docx')).toEqual({
    published:{parsed:1,preserved:1,total:2}, commit:{parsed:2,preserved:1,total:2},
  });
  const text = renderSection(merged).split('### DOCX')[1].split('### PPTX')[0];
  expect(text).toContain('<td>Parse success</td><td align="right">50.00%</td><td align="right">100.00%</td>');
  expect(text).toContain('<td>Lossless roundtrip</td><td align="right">50.00%</td><td align="right">50.00%</td>');
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
  ]) {
    const input = parts();
    mutate(input);
    expect(() => mergeRoundtrips(plan(), plan(), input, hash)).toThrow();
  }
});


test('unmeasured LibreOffice parse/preservation stays blank and stale results are rejected', () => {
  const merged = mergeRoundtrips(plan(), plan(), parts(), hash);
  const report = {...merged,xlsx_fidelity_benchmark:{source_sha:sha,libreoffice_version:'26.2.3.2'}};
  const text = renderSection(report).split('### XLSX')[1];
  expect(text).toContain('<td>Parse success</td><td align="right">100.00%</td><td align="right">100.00%</td><td align="right">—</td>');
  expect(text).toContain('<td>Lossless roundtrip</td><td align="right">100.00%</td><td align="right">100.00%</td><td align="right">—</td>');
  expect(() => renderSection({...merged,source_sha:published})).toThrow('revision');
});
