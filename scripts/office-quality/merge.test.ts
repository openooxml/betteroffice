import { expect, test } from 'bun:test';
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { mergeFromPaths, mergeReports } from './merge.mjs';

const commit = 'c'.repeat(40);
const sourceSha = 'd'.repeat(40);
const versions = { docx: '1.2.3', pptx: '2.3.4', xlsx: '3.4.5' };

function metadata(format: string, hash: string) {
  return {
    format,
    source: { sha256: hash },
    reference: { status: 'ok', dpi: 150, sha256: hash, pages: 1 },
    reference_pages: [
      { url: 'https://corpus.example/page_0001.png', bytes: 1, sha256: hash },
    ],
  };
}

function plan() {
  return {
    schema_version: 1,
    source_sha: sourceSha,
    commit,
    versions,
    react_version: '4.5.6',
    formats: ['docx', 'pptx'],
    samples: [
      { id: 'docx-a', format: 'docx', metadata: metadata('docx', 'a'.repeat(64)) },
      { id: 'docx-b', format: 'docx', metadata: metadata('docx', 'b'.repeat(64)) },
      { id: 'pptx-c', format: 'pptx', metadata: metadata('pptx', 'e'.repeat(64)) },
    ],
  };
}

function successful(
  channel: 'published' | 'commit',
  hash: string,
  value: number,
  referencePages = 1,
  actualPages = referencePages
) {
  return {
    channel,
    ...(channel === 'published'
      ? { version: versions.docx }
      : { renderer_source_commit: commit }),
    status: 'ok',
    source_verified: true,
    reference: { status: 'ok', sha256: hash },
    actual: { status: 'ok', sha256: hash },
    penalized_ssim: value,
    resized: false,
    reference_pages: referencePages,
    actual_pages: actualPages,
  };
}

function sample(
  id: string,
  format: string,
  hash: string,
  published: Record<string, unknown>,
  commitResult: Record<string, unknown>
) {
  return { id, format, source_sha256: hash, comparisons: [published, commitResult] };
}

function reports() {
  const input = plan();
  return [
    {
      source_sha: sourceSha,
      commit,
      versions,
      react_version: '4.5.6',
      format: 'docx',
      samples: [
        sample(
          'docx-a',
          'docx',
          'a'.repeat(64),
          successful('published', 'a'.repeat(64), 0.2),
          successful('commit', 'a'.repeat(64), 0.2, 1, 2)
        ),
        sample(
          'docx-b',
          'docx',
          'b'.repeat(64),
          successful('published', 'b'.repeat(64), 0.8),
          successful('commit', 'b'.repeat(64), 0.8)
        ),
      ],
    },
    {
      source_sha: sourceSha,
      commit,
      versions,
      react_version: '4.5.6',
      format: 'pptx',
      samples: [
        sample(
          'pptx-c',
          'pptx',
          'e'.repeat(64),
          {
            channel: 'published',
            version: versions.pptx,
            status: 'failed',
            stage: 'capture',
            error: 'capture stopped',
          },
          { ...successful('commit', 'e'.repeat(64), 0.6), renderer_source_commit: commit }
        ),
      ],
    },
  ];
}

function fixture(root: string) {
  return {
    plan: join(root, 'plan.json'),
    parts: join(root, 'parts'),
    output: join(root, 'output'),
  };
}

async function writeFixture(root: string, requireRenders = false) {
  const paths = fixture(root);
  await writeFile(paths.plan, JSON.stringify(plan()));
  for (const report of reports()) {
    const directory = join(paths.parts, `visual-fidelity-report-${report.format}`);
    await mkdir(directory, { recursive: true });
    await writeFile(join(directory, 'report.json'), JSON.stringify(report));
  }
  if (requireRenders) {
    for (const [id, pages] of [
      ['docx-a', 2],
      ['docx-b', 1],
      ['pptx-c', 1],
    ] as const) {
      const format = id.startsWith('docx') ? 'docx' : 'pptx';
      const directory = join(
        paths.parts,
        `visual-fidelity-renders-${format}`,
        id,
        'commit'
      );
      await mkdir(directory, { recursive: true });
      for (let page = 1; page <= pages; page += 1)
        await writeFile(
          join(directory, `page_${String(page).padStart(4, '0')}.png`),
          `render ${id} ${page}`
        );
      await writeFile(
        join(paths.parts, `visual-fidelity-renders-${format}`, 'report.json'),
        '{}'
      );
    }
  }
  return paths;
}

test('merges original samples in plan order and retains failed channel coverage', async () => {
  const merged = mergeReports(plan(), reports());
  expect(merged.samples.map((entry) => entry.id)).toEqual(['docx-a', 'docx-b', 'pptx-c']);
  expect(merged.samples[2].comparisons[0]).toMatchObject({
    status: 'failed',
    stage: 'capture',
  });
  const section = (await import('./readme.mjs')).renderSection(merged);
  expect(section).toContain(
    '<tr><td>SSIM</td><td align="right">0.5000</td><td align="right">0.5000</td></tr>'
  );
  expect(section).toContain(
    '<tr><td>SSIM</td><td align="right">—</td><td align="right">0.6000</td></tr>'
  );
});

test('rejects missing, duplicate, mismatched, source-mismatched, and cross-format reports', () => {
  expect(() => mergeReports(plan(), reports().slice(0, 1))).toThrow(
    'missing pptx report'
  );
  expect(() => mergeReports(plan(), [...reports(), reports()[0]])).toThrow('duplicate');
  const badIdentity = reports();
  badIdentity[0].react_version = 'other';
  expect(() => mergeReports(plan(), badIdentity)).toThrow('identity');
  const badSourceIdentity = reports();
  badSourceIdentity[0].source_sha = 'f'.repeat(40);
  expect(() => mergeReports(plan(), badSourceIdentity)).toThrow('identity');
  const badCommitIdentity = reports();
  badCommitIdentity[0].commit = 'f'.repeat(40);
  expect(() => mergeReports(plan(), badCommitIdentity)).toThrow('identity');
  const badVersionIdentity = reports();
  badVersionIdentity[0].versions = { ...versions, docx: '9.9.9' };
  expect(() => mergeReports(plan(), badVersionIdentity)).toThrow('identity');
  const badChannelIdentity = reports();
  badChannelIdentity[0].samples[0].comparisons[0].version = '9.9.9';
  expect(() => mergeReports(plan(), badChannelIdentity)).toThrow('identity');
  const badSource = reports();
  badSource[0].samples[0].source_sha256 = 'f'.repeat(64);
  expect(() => mergeReports(plan(), badSource)).toThrow('source hash');
  const crossFormat = reports();
  crossFormat[0].samples[0].format = 'pptx';
  expect(() => mergeReports(plan(), crossFormat)).toThrow('cross-format');
  const missing = reports();
  missing[0].samples.pop();
  expect(() => mergeReports(plan(), missing)).toThrow('missing samples');
  const duplicateSample = reports();
  duplicateSample[0].samples[1] = structuredClone(duplicateSample[0].samples[0]);
  expect(() => mergeReports(plan(), duplicateSample)).toThrow('unknown or duplicate');
  const missingChannel = reports();
  missingChannel[0].samples[0].comparisons[1] = structuredClone(
    missingChannel[0].samples[0].comparisons[0]
  );
  expect(() => mergeReports(plan(), missingChannel)).toThrow(
    'one published and one commit'
  );
  const wrongReferenceHash = reports();
  wrongReferenceHash[0].samples[0].comparisons[0].reference.sha256 = 'f'.repeat(64);
  wrongReferenceHash[0].samples[0].comparisons[0].actual.sha256 = 'f'.repeat(64);
  expect(() => mergeReports(plan(), wrongReferenceHash)).toThrow('reference hash');
  const wrongReferencePages = reports();
  wrongReferencePages[0].samples[0].comparisons[0].reference_pages = 2;
  expect(() => mergeReports(plan(), wrongReferencePages)).toThrow('reference pages');
  const unexpectedStatus = reports();
  unexpectedStatus[0].samples[0].comparisons[0].status = 'unknown';
  expect(() => mergeReports(plan(), unexpectedStatus)).toThrow('invalid successful');
});

test('writes only expected successful commit pages when render collection is required', async () => {
  const root = await mkdtemp(join(tmpdir(), 'fidelity-merge-'));
  try {
    const paths = await writeFixture(root, true);
    await mergeFromPaths({
      planPath: paths.plan,
      parts: paths.parts,
      output: paths.output,
      requireRenders: true,
    });
    expect(
      JSON.parse(await readFile(join(paths.output, 'report.json'), 'utf8')).samples
    ).toHaveLength(3);
    expect(
      await readFile(join(paths.output, 'docx-a', 'commit', 'page_0002.png'), 'utf8')
    ).toBe('render docx-a 2');
    expect(
      await readFile(join(paths.output, 'docx-b', 'commit', 'page_0001.png'), 'utf8')
    ).toBe('render docx-b 1');
    expect(
      await readFile(join(paths.output, 'pptx-c', 'commit', 'page_0001.png'), 'utf8')
    ).toBe('render pptx-c 1');
    expect(await Bun.file(join(paths.output, 'parts', 'report.json')).exists()).toBe(
      false
    );
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test('skips a successful zero-page commit render', async () => {
  const root = await mkdtemp(join(tmpdir(), 'fidelity-merge-'));
  try {
    const paths = await writeFixture(root, true);
    const pptx = reports()[1];
    pptx.samples[0].comparisons[1].actual_pages = 0;
    pptx.samples[0].comparisons[1].penalized_ssim = 0;
    await writeFile(
      join(paths.parts, 'visual-fidelity-report-pptx', 'report.json'),
      JSON.stringify(pptx)
    );
    await rm(join(paths.parts, 'visual-fidelity-renders-pptx'), { recursive: true });
    await mergeFromPaths({
      planPath: paths.plan,
      parts: paths.parts,
      output: paths.output,
      requireRenders: true,
    });
    expect(
      await Bun.file(join(paths.output, 'pptx-c', 'commit', 'page_0001.png')).exists()
    ).toBe(false);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test('rejects missing and symlinked render pages without publishing a partial result', async () => {
  const root = await mkdtemp(join(tmpdir(), 'fidelity-merge-'));
  const unsafe = await mkdtemp(join(tmpdir(), 'fidelity-merge-'));
  try {
    const paths = await writeFixture(root, true);
    await Bun.write(
      join(
        paths.parts,
        'visual-fidelity-renders-docx',
        'docx-a',
        'commit',
        'page_0002.png'
      ),
      ''
    );
    await expect(
      mergeFromPaths({
        planPath: paths.plan,
        parts: paths.parts,
        output: paths.output,
        requireRenders: true,
      })
    ).rejects.toThrow('missing, empty, or unsafe');
    expect(await Bun.file(join(paths.output, 'report.json')).exists()).toBe(false);

    const unsafePaths = await writeFixture(unsafe, true);
    const render = join(
      unsafePaths.parts,
      'visual-fidelity-renders-docx',
      'docx-a',
      'commit',
      'page_0002.png'
    );
    await Bun.write(render, 'target');
    const link = join(
      unsafePaths.parts,
      'visual-fidelity-renders-docx',
      'docx-a',
      'commit',
      'page_0001.png'
    );
    await rm(link);
    await symlink(render, link);
    await expect(
      mergeFromPaths({
        planPath: unsafePaths.plan,
        parts: unsafePaths.parts,
        output: unsafePaths.output,
        requireRenders: true,
      })
    ).rejects.toThrow('unsafe');
  } finally {
    await Promise.all([
      rm(root, { recursive: true, force: true }),
      rm(unsafe, { recursive: true, force: true }),
    ]);
  }
});
